mod docx;
mod error;
mod fonts;
mod geometry;
mod math;
mod model;
mod pdf;

pub use error::Error;

use std::path::{Path, PathBuf};
use std::time::Instant;

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
}

impl Default for ConvertOptions {
    fn default() -> Self {
        Self {
            revisions: RevisionView::Final,
            resample_pictures: true,
            font_dirs: Vec::new(),
            font_cache: true,
            debug_wrap: false,
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

    pub(crate) fn font_setup(&self) -> fonts::FontSetup {
        fonts::FontSetup {
            dirs: self.font_dirs.clone(),
            cache: self.font_cache,
        }
    }
}

/// Converts with [`ConvertOptions::from_env`].
pub fn convert_docx_to_pdf(input: impl AsRef<Path>, path: impl AsRef<Path>) -> Result<(), Error> {
    convert_docx_to_pdf_with(input, path, &ConvertOptions::from_env())
}

pub fn convert_docx_to_pdf_with(
    input: impl AsRef<Path>,
    path: impl AsRef<Path>,
    options: &ConvertOptions,
) -> Result<(), Error> {
    reject_old_env_vars()?;
    let doc = docx::parse(input.as_ref(), options.revisions == RevisionView::Markup)?;
    render_and_write(&doc, path, options)
}

/// Converts with [`ConvertOptions::from_env`].
pub fn convert_docx_bytes_to_pdf(input: &[u8], path: impl AsRef<Path>) -> Result<(), Error> {
    convert_docx_bytes_to_pdf_with(input, path, &ConvertOptions::from_env())
}

pub fn convert_docx_bytes_to_pdf_with(
    input: &[u8],
    path: impl AsRef<Path>,
    options: &ConvertOptions,
) -> Result<(), Error> {
    reject_old_env_vars()?;
    let doc = docx::parse_bytes(input, options.revisions == RevisionView::Markup)?;
    render_and_write(&doc, path, options)
}

/// For callers that serve or store the PDF themselves (web handlers, databases,
/// wasm). Converts with [`ConvertOptions::from_env`].
pub fn convert_docx_bytes_to_pdf_bytes(input: &[u8]) -> Result<Vec<u8>, Error> {
    convert_docx_bytes_to_pdf_bytes_with(input, &ConvertOptions::from_env())
}

pub fn convert_docx_bytes_to_pdf_bytes_with(
    input: &[u8],
    options: &ConvertOptions,
) -> Result<Vec<u8>, Error> {
    reject_old_env_vars()?;
    let doc = docx::parse_bytes(input, options.revisions == RevisionView::Markup)?;
    pdf::render(&doc, options)
}

/// The variables lost their old `DOCXSIDE_` prefix; failing loudly beats silently
/// dropping someone's font directories and rendering with fallbacks.
fn reject_old_env_vars() -> Result<(), Error> {
    match old_env_var(std::env::vars_os().map(|(key, _)| key)) {
        Some(name) => Err(Error::RenamedEnvVar(name)),
        None => Ok(()),
    }
}

fn old_env_var(keys: impl Iterator<Item = std::ffi::OsString>) -> Option<String> {
    keys.filter_map(|key| key.into_string().ok())
        .find(|key| key.starts_with("DOCXSIDE_"))
}

fn render_and_write(
    doc: &model::Document,
    path: impl AsRef<Path>,
    options: &ConvertOptions,
) -> Result<(), Error> {
    let path = path.as_ref().with_extension("pdf");
    let t0 = Instant::now();

    let bytes = pdf::render(doc, options)?;
    let t_render = t0.elapsed();

    std::fs::write(&path, &bytes)?;
    let t_total = t0.elapsed();

    log::info!(
        "Timing: render={:.1}ms, write={:.1}ms, total={:.1}ms (output {} bytes)",
        t_render.as_secs_f64() * 1000.0,
        (t_total - t_render).as_secs_f64() * 1000.0,
        t_total.as_secs_f64() * 1000.0,
        bytes.len(),
    );

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn old_prefix_is_reported_with_its_new_name() {
        let keys = ["PATH", "DOCXIDE_FONTS", "DOCXSIDE_FONTS"].map(Into::into);
        let old = old_env_var(keys.into_iter()).unwrap();
        assert_eq!(
            Error::RenamedEnvVar(old).to_string(),
            "DOCXSIDE_FONTS is no longer read; rename it to DOCXIDE_FONTS"
        );
        assert_eq!(old_env_var(["DOCXIDE_FONTS".into()].into_iter()), None);
    }

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
