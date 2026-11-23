//! Converts DOCX to PDF, matching Microsoft Word's PDF export (the online
//! export: tagged, Windows fonts and metrics).
//!
//! [`convert`] takes the DOCX as bytes and returns a [`Conversion`]: the PDF
//! as bytes, where bookmarked paragraphs were laid out, the fonts worth
//! supplying, and optionally the pages as a [`display`] list. File I/O is the
//! caller's. [`ConvertOptions::new`] reproduces Word's online export.
//!
//! ```no_run
//! use docxide_pdf::{ConvertOptions, convert};
//!
//! let out = convert(&std::fs::read("input.docx")?, &ConvertOptions::new())?;
//! std::fs::write("output.pdf", &out.pdf)?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
#![warn(missing_docs)]

mod compression;
pub mod display;
mod docx;
mod error;
mod fast_hash;
mod fonts;
mod geometry;
mod math;
mod model;
mod pdf;

pub use compression::Compression;
pub use error::Error;
pub use fonts::missing::MissingFont;
pub use pdf::bookmarks::BookmarkPosition;

use fonts::missing::take_missing_fonts;
use pdf::bookmarks::take_bookmark_positions;
use std::path::PathBuf;

/// Which view of a document's tracked changes to print.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum RevisionView {
    /// Word's "No Markup": insertions plain, deletions dropped. What Word's
    /// online export prints.
    #[default]
    Final,
    /// Word's "All Markup", inline: insertions underlined and deletions
    /// struck through in the author's colour, deleted text kept in the
    /// layout, a change bar in the margin beside every revised line.
    Markup,
}

/// Invisible link annotations over the text laid out between two bookmarks:
/// from `start` (a `w:bookmarkStart` name) to `end` (another one), wherever
/// they sit — between paragraphs or tables, or between the runs of a line.
///
/// Every line holding text (or a picture) of the range gets the link, as the
/// box of its glyphs in that range; per page, the range's first and last
/// lines keep their own rectangle and the lines in between share one, so a
/// range has at most three rectangles a page. Only the document body is
/// covered: headers, footers and notes are not.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RangeLink {
    /// The bookmark the range starts at.
    pub start: String,
    /// The bookmark the range ends at; the text before it is in the range.
    pub end: String,
    /// The link's URI action.
    pub uri: String,
    /// A second link over the range's heading, when it has one: its leading
    /// lines on the first page it reaches, while every run on them is bold
    /// (at most four lines, without a gap taller than a line between them).
    pub heading_uri: Option<String>,
}

impl RangeLink {
    /// Link the text between the named start and end bookmarks.
    pub fn new(start: impl Into<String>, end: impl Into<String>, uri: impl Into<String>) -> Self {
        Self {
            start: start.into(),
            end: end.into(),
            uri: uri.into(),
            heading_uri: None,
        }
    }

    /// Add a separate URI over the range's leading bold lines.
    pub fn heading_uri(mut self, uri: impl Into<String>) -> Self {
        self.heading_uri = Some(uri.into());
        self
    }
}

/// `START END URI [HEADING_URI]`, separated by whitespace (which neither a
/// bookmark name nor a URI contains): the spelling of `DOCXIDE_RANGE_LINKS`
/// lines and of `--range-link`.
impl std::str::FromStr for RangeLink {
    type Err = String;

    fn from_str(text: &str) -> Result<Self, Self::Err> {
        let fields: Vec<&str> = text.split_whitespace().collect();
        match fields[..] {
            [start, end, uri] => Ok(Self::new(start, end, uri)),
            [start, end, uri, heading] => Ok(Self::new(start, end, uri).heading_uri(heading)),
            _ => Err(format!(
                "range link {text:?}: expected START END URI [HEADING_URI]"
            )),
        }
    }
}

/// Per-conversion settings, set with the builder methods. The defaults
/// reproduce Word's online export; each option moves one behaviour towards
/// desktop Word or adapts the conversion to its host. The library reads no
/// environment variables (the CLI maps `DOCXIDE_*` onto these options).
#[derive(Clone, Debug)]
#[non_exhaustive]
pub struct ConvertOptions {
    revisions: RevisionView,
    resample_pictures: bool,
    font_dirs: Vec<PathBuf>,
    font_cache: bool,
    debug_wrap: bool,
    display_list: bool,
    compression: Compression,
    range_links: Vec<RangeLink>,
    fonts: std::sync::Arc<fonts::memory::MemoryFonts>,
    system_fonts: bool,
}

impl Default for ConvertOptions {
    fn default() -> Self {
        Self {
            revisions: RevisionView::Final,
            resample_pictures: true,
            font_dirs: Vec::new(),
            font_cache: true,
            debug_wrap: false,
            display_list: false,
            compression: Compression::Smallest,
            range_links: Vec::new(),
            fonts: Default::default(),
            // A bare wasm instance has no font directories to scan.
            system_fonts: !cfg!(all(target_arch = "wasm32", target_os = "unknown")),
        }
    }
}

impl ConvertOptions {
    /// Word's online export behaviour.
    pub fn new() -> Self {
        Self::default()
    }

    /// Which view of the tracked changes to print.
    pub fn revisions(mut self, view: RevisionView) -> Self {
        self.revisions = view;
        self
    }

    /// Resample pictures shown above 192 ppi down to 192 ppi, as Word's online
    /// export does (the default). `false` keeps the source pixels, like a
    /// desktop export with picture compression off.
    pub fn resample_pictures(mut self, on: bool) -> Self {
        self.resample_pictures = on;
        self
    }

    /// Font directories searched before the system ones; the first listed
    /// wins.
    pub fn font_dirs(mut self, dirs: impl IntoIterator<Item = impl Into<PathBuf>>) -> Self {
        self.font_dirs = dirs.into_iter().map(Into::into).collect();
        self
    }

    /// A TrueType/OpenType font (all faces of a collection) to use before
    /// installed faces of the same family and style. Parsed here, once:
    /// reusing the options re-parses nothing. Data without a usable face is
    /// skipped with a warning in the log.
    pub fn font(mut self, data: Vec<u8>) -> Self {
        if std::sync::Arc::make_mut(&mut self.fonts).add(data) == 0 {
            log::warn!("ConvertOptions::font: no usable font face in the data");
        }
        self
    }

    /// Search the system font directories (the default, except on bare
    /// wasm). `false` keeps the `font_dirs`, the `font`s and the fonts the
    /// document embeds.
    pub fn system_fonts(mut self, on: bool) -> Self {
        self.system_fonts = on;
        self
    }

    /// Read and write the font index disk cache (on by default).
    pub fn font_cache(mut self, on: bool) -> Self {
        self.font_cache = on;
        self
    }

    /// Draw the float wrap zones on the page: a layout debugging aid for the
    /// CLI, not part of the stable API.
    #[doc(hidden)]
    pub fn debug_wrap(mut self, on: bool) -> Self {
        self.debug_wrap = on;
        self
    }

    /// Also return the pages as display items ([`Conversion::display`]),
    /// recorded from the same render as the PDF.
    pub fn display_list(mut self, on: bool) -> Self {
        self.display_list = on;
        self
    }

    /// How hard to compress PDF streams. The default keeps the smallest file;
    /// [`Compression::Fast`] trades size for conversion time.
    pub fn compression(mut self, compression: Compression) -> Self {
        self.compression = compression;
        self
    }

    /// Write invisible links over text between bookmark pairs.
    pub fn range_links(mut self, links: impl IntoIterator<Item = RangeLink>) -> Self {
        self.range_links = links.into_iter().collect();
        self
    }

    pub(crate) fn font_setup(&self) -> fonts::FontSetup {
        fonts::FontSetup {
            dirs: self.font_dirs.clone(),
            cache: self.font_cache,
            system: self.system_fonts,
        }
    }
}

/// Everything one conversion produces.
#[non_exhaustive]
#[derive(Debug, Clone)]
pub struct Conversion {
    /// The PDF file.
    pub pdf: Vec<u8>,
    /// Where each bookmarked body or table-cell paragraph was laid out.
    pub bookmarks: Vec<BookmarkPosition>,
    /// The fonts the document asks for that no face provides.
    pub missing_fonts: Vec<MissingFont>,
    /// What each page draws, with the fonts and pictures it draws them with,
    /// for a caller that writes PDF (or renders pages) itself. `Some` only
    /// with [`ConvertOptions::display_list`].
    pub display: Option<display::Document>,
}

/// Converts DOCX bytes. File I/O is the caller's:
/// `std::fs::write("out.pdf", &convert(&std::fs::read("in.docx")?, &options)?.pdf)`.
pub fn convert(docx: &[u8], options: &ConvertOptions) -> Result<Conversion, Error> {
    // ponytail: bookmarks and missing fonts are thread-locals drained here, so a
    // conversion nested inside another on one thread would mix them; nothing nests.
    take_bookmark_positions();
    take_missing_fonts();
    let _fonts = fonts::memory::install(&options.fonts);
    let _compression = compression::scoped(options.compression);
    let recording = options.display_list.then(pdf::canvas::Recording::start);
    let pdf = docx::parse_bytes_with(docx, options).and_then(|doc| pdf::render(&doc, options));
    let bookmarks = take_bookmark_positions();
    let missing_fonts = take_missing_fonts();
    Ok(Conversion {
        pdf: pdf?,
        bookmarks,
        missing_fonts,
        display: recording.map(|r| r.finish().unwrap_or_default()),
    })
}

// A server converts on worker threads; an `Rc` in either type must not slip in.
const _: fn() = || {
    fn send_sync<T: Send + Sync>() {}
    send_sync::<ConvertOptions>();
    send_sync::<Conversion>();
};

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builders_override_the_online_defaults() {
        let plain = ConvertOptions::new();
        assert_eq!(plain.revisions, RevisionView::Final);
        assert!(plain.resample_pictures && plain.font_cache && !plain.debug_wrap);
        assert!(plain.font_dirs.is_empty() && !plain.display_list);
        assert_eq!(plain.compression, Compression::Smallest);
        let built = ConvertOptions::new()
            .revisions(RevisionView::Markup)
            .resample_pictures(false)
            .font_dirs(["x"])
            .font_cache(false)
            .debug_wrap(true)
            .display_list(true)
            .compression(Compression::Fast);
        assert_eq!(built.revisions, RevisionView::Markup);
        assert_eq!(built.compression, Compression::Fast);
        assert!(!built.resample_pictures && !built.font_cache && built.debug_wrap);
        assert!(built.display_list);
        assert_eq!(built.font_setup().dirs, [PathBuf::from("x")]);
    }

    #[test]
    fn range_links_read_from_lines_of_start_end_uri_and_heading() {
        let link = |vars: &str| {
            vars.lines()
                .filter_map(|line| line.parse::<RangeLink>().ok())
                .collect::<Vec<_>>()
        };
        assert!(ConvertOptions::new().range_links.is_empty());
        assert_eq!(
            link("a b x:1\n\n  c  d x:2 x:h \nbad line\ne f g h i"),
            [
                RangeLink::new("a", "b", "x:1"),
                RangeLink::new("c", "d", "x:2").heading_uri("x:h"),
            ]
        );
        assert!("a b".parse::<RangeLink>().is_err());
        let built = ConvertOptions::new().range_links([RangeLink::new("a", "b", "x:1")]);
        assert_eq!(built.range_links.len(), 1);
    }
}
