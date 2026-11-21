use std::fmt;

#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
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
