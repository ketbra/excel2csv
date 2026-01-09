use crate::error::{Error, Result};
use std::path::Path;
use umya_spreadsheet::Spreadsheet;

pub fn open_workbook(path: &Path) -> Result<Spreadsheet> {
    if !path.exists() {
        return Err(Error::FileNotFound(path.to_path_buf()));
    }

    let extension = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_default();

    match extension.as_str() {
        "xlsx" | "xlsm" => open_xlsx(path),
        "xls" => open_xls(path),
        _ => Err(Error::UnsupportedFormat(extension)),
    }
}

fn open_xlsx(path: &Path) -> Result<Spreadsheet> {
    umya_spreadsheet::reader::xlsx::read(path).map_err(|e| Error::InvalidExcel {
        path: path.to_path_buf(),
        details: e.to_string(),
    })
}

fn open_xls(path: &Path) -> Result<Spreadsheet> {
    xlrd::open(path).map_err(|e| Error::InvalidExcel {
        path: path.to_path_buf(),
        details: e.to_string(),
    })
}
