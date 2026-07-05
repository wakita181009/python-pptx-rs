use thiserror::Error;

#[derive(Debug, Error)]
pub enum Error {
    #[error("package error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
    #[error("XML parse error: {0}")]
    Xml(String),
    #[error("part not found: {0}")]
    PartNotFound(String),
    #[error("invalid package: {0}")]
    InvalidPackage(String),
}

pub type Result<T> = std::result::Result<T, Error>;
