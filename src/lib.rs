mod docx;
mod error;
mod model;
mod pdf;

pub use error::Error;

use std::path::Path;

pub fn convert_docx_to_pdf(input: &Path, output: &Path) -> Result<(), Error> {
    let doc = docx::parse(input)?;
    let bytes = pdf::render(&doc)?;
    std::fs::write(output, bytes).map_err(Error::Io)
}
