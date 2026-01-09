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

/// Get list of sheet names from workbook
pub fn get_sheet_names(workbook: &Spreadsheet) -> Vec<String> {
    let mut names = Vec::new();
    let mut index = 0usize;
    while let Some(sheet) = workbook.get_sheet(&index) {
        names.push(sheet.get_name().to_string());
        index += 1;
    }
    names
}

/// Get sheet count
pub fn get_sheet_count(workbook: &Spreadsheet) -> usize {
    let mut count = 0usize;
    while workbook.get_sheet(&count).is_some() {
        count += 1;
    }
    count
}

/// Parse sheet selector (name or index)
pub fn parse_sheet_selector(selector: &str) -> SheetSelector {
    if let Ok(index) = selector.parse::<usize>() {
        SheetSelector::Index(index)
    } else {
        SheetSelector::Name(selector.to_string())
    }
}

#[derive(Debug, Clone)]
pub enum SheetSelector {
    Index(usize),
    Name(String),
}
