pub mod display;
mod docx;
mod error;
mod fonts;
mod geometry;
mod math;
mod model;
mod pdf;

pub use error::Error;
pub use fonts::memory::{clear_fonts, register_font, use_system_fonts};
pub use fonts::requests::MissingFont;
pub use pdf::bookmarks::BookmarkPosition;

use fonts::requests::take_font_requests;
use pdf::bookmarks::take_bookmark_positions;
use std::path::PathBuf;

/// Which view of a document's tracked changes to print.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
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

/// Per-conversion settings. The defaults reproduce Word's online export; each
/// option moves one behaviour towards desktop Word or adapts the conversion to
/// its host. Explicit options never read the environment;
/// [`ConvertOptions::from_env`] reads the same settings from `DOCXIDE_*`
/// variables and is what the functions without an options argument use.
#[derive(Clone, Debug)]
#[non_exhaustive]
pub struct ConvertOptions {
    /// Which view of the tracked changes to print. Environment:
    /// `DOCXIDE_REVISIONS=markup`.
    pub revisions: RevisionView,
    /// Resample pictures shown above 192 ppi down to 192 ppi, as Word's online
    /// export does. `false` keeps the source pixels, like a desktop export
    /// with picture compression off. Environment: `DOCXIDE_NO_RESAMPLE=1`.
    pub resample_pictures: bool,
    /// Font directories searched before the system ones; the first listed
    /// wins. Environment: `DOCXIDE_FONTS`, colon-separated (`;` on Windows).
    pub font_dirs: Vec<PathBuf>,
    /// Read and write the font index disk cache. Environment:
    /// `DOCXIDE_NO_FONT_CACHE=1` switches it off.
    pub font_cache: bool,
    /// Draw the float wrap zones on the page. Environment:
    /// `DOCXIDE_DEBUG_WRAP=1`.
    pub debug_wrap: bool,
    display_list: bool,
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
        }
    }
}

impl ConvertOptions {
    /// Word's online export behaviour.
    pub fn new() -> Self {
        Self::default()
    }

    /// The defaults overridden by whichever `DOCXIDE_*` variables are set. A
    /// boolean variable is on unless its value is empty, `0` or `false`.
    pub fn from_env() -> Self {
        Self::from_vars(|name| std::env::var(name).ok())
    }

    fn from_vars(var: impl Fn(&str) -> Option<String>) -> Self {
        let flag = |name: &str| {
            var(name).is_some_and(|v| {
                !matches!(v.trim().to_ascii_lowercase().as_str(), "" | "0" | "false")
            })
        };
        let mut options = Self::new();
        if var("DOCXIDE_REVISIONS").is_some_and(|v| v.trim().eq_ignore_ascii_case("markup")) {
            options.revisions = RevisionView::Markup;
        }
        if flag("DOCXIDE_NO_RESAMPLE") {
            options.resample_pictures = false;
        }
        if let Some(dirs) = var("DOCXIDE_FONTS") {
            let sep = if cfg!(windows) { ';' } else { ':' };
            options.font_dirs = dirs
                .split(sep)
                .map(str::trim)
                .filter(|dir| !dir.is_empty())
                .map(PathBuf::from)
                .collect();
        }
        if flag("DOCXIDE_NO_FONT_CACHE") {
            options.font_cache = false;
        }
        if flag("DOCXIDE_DEBUG_WRAP") {
            options.debug_wrap = true;
        }
        options
    }

    pub fn revisions(mut self, view: RevisionView) -> Self {
        self.revisions = view;
        self
    }

    pub fn resample_pictures(mut self, on: bool) -> Self {
        self.resample_pictures = on;
        self
    }

    pub fn font_dirs(mut self, dirs: impl IntoIterator<Item = impl Into<PathBuf>>) -> Self {
        self.font_dirs = dirs.into_iter().map(Into::into).collect();
        self
    }

    pub fn font_cache(mut self, on: bool) -> Self {
        self.font_cache = on;
        self
    }

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

    pub(crate) fn font_setup(&self) -> fonts::FontSetup {
        fonts::FontSetup {
            dirs: self.font_dirs.clone(),
            cache: self.font_cache,
        }
    }
}

/// Everything one conversion produces.
#[non_exhaustive]
#[derive(Debug, Clone)]
pub struct Conversion {
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
    take_font_requests();
    let recording = options.display_list.then(pdf::canvas::Recording::start);
    let pdf = docx::parse_bytes(docx, options.revisions).and_then(|doc| pdf::render(&doc, options));
    let bookmarks = take_bookmark_positions();
    let missing_fonts = take_font_requests();
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
    fn env_vars_override_the_online_defaults() {
        let with = |vars: &[(&str, &str)]| {
            ConvertOptions::from_vars(|name| {
                vars.iter()
                    .find(|(n, _)| *n == name)
                    .map(|(_, v)| v.to_string())
            })
        };
        let plain = with(&[]);
        assert_eq!(plain.revisions, RevisionView::Final);
        assert!(plain.resample_pictures && plain.font_cache && !plain.debug_wrap);
        assert!(plain.font_dirs.is_empty());
        for off in ["0", "false", "", " False "] {
            assert!(with(&[("DOCXIDE_NO_RESAMPLE", off)]).resample_pictures);
        }
        for not_markup in ["", "final", "1"] {
            assert_eq!(
                with(&[("DOCXIDE_REVISIONS", not_markup)]).revisions,
                RevisionView::Final
            );
        }
        let fonts = if cfg!(windows) {
            " a ; ;b;c "
        } else {
            " a : :b:c "
        };
        let set = with(&[
            ("DOCXIDE_REVISIONS", " Markup "),
            ("DOCXIDE_NO_RESAMPLE", "1"),
            ("DOCXIDE_NO_FONT_CACHE", "yes"),
            ("DOCXIDE_DEBUG_WRAP", "true"),
            ("DOCXIDE_FONTS", fonts),
        ]);
        assert_eq!(set.revisions, RevisionView::Markup);
        assert!(!set.resample_pictures && !set.font_cache && set.debug_wrap);
        assert_eq!(set.font_dirs, [PathBuf::from("a"), "b".into(), "c".into()]);
        let built = ConvertOptions::new()
            .revisions(RevisionView::Markup)
            .resample_pictures(false)
            .font_dirs(["x"])
            .font_cache(false)
            .debug_wrap(true);
        assert_eq!(built.revisions, RevisionView::Markup);
        assert!(!built.resample_pictures && !built.font_cache && built.debug_wrap);
        assert_eq!(built.font_setup().dirs, [PathBuf::from("x")]);
    }
}
