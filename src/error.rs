use std::fmt;

/// Why [`crate::convert`] failed.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The input is not a DOCX: not a ZIP archive, or no readable
    /// `word/document.xml` with a body. The text says which.
    InvalidDocx(String),
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::InvalidDocx(reason) => write!(f, "not a valid DOCX file: {reason}"),
        }
    }
}

impl std::error::Error for Error {}
