mod docx;
mod error;
mod fonts;
mod geometry;
mod model;
mod pdf;

pub use error::Error;

use std::path::Path;
use std::time::Instant;

pub fn convert_docx_to_pdf(input: impl AsRef<Path>, path: impl AsRef<Path>) -> Result<(), Error> {
    let doc = docx::parse(input.as_ref())?;
    render_and_write(&doc, path)
}

pub fn convert_docx_bytes_to_pdf(input: &[u8], path: impl AsRef<Path>) -> Result<(), Error> {
    let doc = docx::parse_bytes(input)?;
    render_and_write(&doc, path)
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
