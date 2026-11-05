mod docx;
mod error;
mod fonts;
mod geometry;
mod math;
mod model;
mod pdf;

pub use error::Error;

use std::path::Path;
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

#[derive(Clone, Debug, Default)]
#[non_exhaustive]
pub struct ConvertOptions {
    pub revisions: RevisionView,
}

impl ConvertOptions {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn revisions(mut self, view: RevisionView) -> Self {
        self.revisions = view;
        self
    }
}

pub fn convert_docx_to_pdf(input: impl AsRef<Path>, path: impl AsRef<Path>) -> Result<(), Error> {
    convert_docx_to_pdf_with(input, path, &ConvertOptions::default())
}

pub fn convert_docx_to_pdf_with(
    input: impl AsRef<Path>,
    path: impl AsRef<Path>,
    options: &ConvertOptions,
) -> Result<(), Error> {
    reject_old_env_vars()?;
    docx::set_markup(options.revisions == RevisionView::Markup);
    let doc = docx::parse(input.as_ref())?;
    render_and_write(&doc, path)
}

pub fn convert_docx_bytes_to_pdf(input: &[u8], path: impl AsRef<Path>) -> Result<(), Error> {
    convert_docx_bytes_to_pdf_with(input, path, &ConvertOptions::default())
}

pub fn convert_docx_bytes_to_pdf_with(
    input: &[u8],
    path: impl AsRef<Path>,
    options: &ConvertOptions,
) -> Result<(), Error> {
    reject_old_env_vars()?;
    docx::set_markup(options.revisions == RevisionView::Markup);
    let doc = docx::parse_bytes(input)?;
    render_and_write(&doc, path)
}

/// For callers that serve or store the PDF themselves (web handlers, databases, wasm).
pub fn convert_docx_bytes_to_pdf_bytes(input: &[u8]) -> Result<Vec<u8>, Error> {
    convert_docx_bytes_to_pdf_bytes_with(input, &ConvertOptions::default())
}

pub fn convert_docx_bytes_to_pdf_bytes_with(
    input: &[u8],
    options: &ConvertOptions,
) -> Result<Vec<u8>, Error> {
    reject_old_env_vars()?;
    docx::set_markup(options.revisions == RevisionView::Markup);
    pdf::render(&docx::parse_bytes(input)?)
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

fn render_and_write(doc: &model::Document, path: impl AsRef<Path>) -> Result<(), Error> {
    let path = path.as_ref().with_extension("pdf");
    let t0 = Instant::now();

    let bytes = pdf::render(doc)?;
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
}
