mod docx;
mod error;
mod fonts;
mod model;
mod pdf;

pub use error::Error;

use std::path::Path;
use std::time::Instant;

pub fn convert_docx_to_pdf(input: &Path, output: &Path) -> Result<(), Error> {
    let t0 = Instant::now();

    let doc = docx::parse(input)?;
    let t_parse = t0.elapsed();

    let bytes = pdf::render(&doc)?;
    let t_render = t0.elapsed();

    std::fs::write(output, &bytes).map_err(Error::Io)?;
    let t_total = t0.elapsed();

    log::info!(
        "Timing: parse={:.1}ms, render={:.1}ms, write={:.1}ms, total={:.1}ms (output {} bytes)",
        t_parse.as_secs_f64() * 1000.0,
        (t_render - t_parse).as_secs_f64() * 1000.0,
        (t_total - t_render).as_secs_f64() * 1000.0,
        t_total.as_secs_f64() * 1000.0,
        bytes.len(),
    );

    Ok(())
}

pub fn convert_docx_bytes_to_pdf(input: &[u8], output: &Path) -> Result<(), Error> {
    let t0 = Instant::now();

    let doc = docx::parse_bytes(input)?;
    let t_parse = t0.elapsed();

    let bytes = pdf::render(&doc)?;
    let t_render = t0.elapsed();

    std::fs::write(output, &bytes).map_err(Error::Io)?;
    let t_total = t0.elapsed();

    log::info!(
        "Timing: parse={:.1}ms, render={:.1}ms, write={:.1}ms, total={:.1}ms (output {} bytes)",
        t_parse.as_secs_f64() * 1000.0,
        (t_render - t_parse).as_secs_f64() * 1000.0,
        (t_total - t_render).as_secs_f64() * 1000.0,
        t_total.as_secs_f64() * 1000.0,
        bytes.len(),
    );

    Ok(())
}
