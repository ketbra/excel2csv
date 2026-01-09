use std::path::PathBuf;
use thiserror::Error;

#[derive(Error, Debug)]
pub enum Error {
    #[error("file not found: {0}")]
    FileNotFound(PathBuf),

    #[error("invalid Excel file: {path} ({details})")]
    InvalidExcel { path: PathBuf, details: String },

    #[error("sheet \"{name}\" not found (available: {available})")]
    SheetNotFound { name: String, available: String },

    #[error("sheet index {index} out of range (have {count} sheets)")]
    SheetIndexOutOfRange { index: usize, count: usize },

    #[error("multiple sheets require -o <directory>")]
    MultipleSheetNoOutput,

    #[error("unsupported file format: {0}")]
    UnsupportedFormat(String),

    #[error("failed to write CSV: {0}")]
    CsvWrite(#[from] csv::Error),

    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
}

impl Error {
    pub fn exit_code(&self) -> i32 {
        match self {
            Error::FileNotFound(_) => 1,
            Error::InvalidExcel { .. } => 2,
            Error::SheetNotFound { .. } => 3,
            Error::SheetIndexOutOfRange { .. } => 3,
            Error::MultipleSheetNoOutput => 3,
            Error::UnsupportedFormat(_) => 3,
            Error::CsvWrite(_) => 4,
            Error::Io(_) => 1,
        }
    }
}

pub type Result<T> = std::result::Result<T, Error>;
