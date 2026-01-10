//! Formula evaluation module using formualizer-workbook.
//!
//! This module provides formula recalculation capabilities for Excel files,
//! allowing cell values to be recomputed based on their formulas.

use crate::error::{Error, Result};
use chrono::NaiveDate;
use formualizer_workbook::{
    backends::UmyaAdapter,
    traits::{LoadStrategy, SpreadsheetReader},
    Workbook, WorkbookConfig,
};
use std::collections::HashMap;
use std::path::Path;

/// Configuration for formula evaluation.
#[derive(Debug, Clone, Default)]
pub struct EvalConfig {
    /// Cell value overrides to apply before evaluation.
    /// Key format: "SheetName!A1" or "A1" (uses first sheet).
    pub overrides: HashMap<String, String>,

    /// Override for TODAY() function (for reproducible output).
    pub today: Option<NaiveDate>,
}

/// Evaluated workbook containing recalculated cell values.
pub struct EvaluatedWorkbook {
    workbook: Workbook,
    sheet_names: Vec<String>,
}

impl EvaluatedWorkbook {
    /// Load and evaluate an Excel file.
    pub fn load(path: &Path, config: &EvalConfig) -> Result<Self> {
        // Open with UmyaAdapter
        let adapter = UmyaAdapter::open_path(path).map_err(|e| Error::InvalidExcel {
            path: path.to_path_buf(),
            details: format!("formualizer: {}", e),
        })?;

        // Get sheet names before consuming adapter
        let sheet_names = adapter.sheet_names().map_err(|e| Error::InvalidExcel {
            path: path.to_path_buf(),
            details: format!("formualizer: {}", e),
        })?;

        // Create workbook from adapter
        let wb_config = WorkbookConfig::ephemeral();
        let mut workbook =
            Workbook::from_reader(adapter, LoadStrategy::EagerAll, wb_config).map_err(|e| {
                Error::InvalidExcel {
                    path: path.to_path_buf(),
                    details: format!("formualizer: {}", e),
                }
            })?;

        // Apply cell overrides
        for (cell_ref, value) in &config.overrides {
            apply_override(&mut workbook, cell_ref, value, &sheet_names)?;
        }

        // TODO: Handle config.today override for TODAY()/NOW() functions
        // This would require setting a custom date context in formualizer

        // Evaluate all formulas
        workbook.evaluate_all().map_err(|e| Error::InvalidExcel {
            path: path.to_path_buf(),
            details: format!("evaluation failed: {}", e),
        })?;

        Ok(Self {
            workbook,
            sheet_names,
        })
    }

    /// Get the evaluated value of a cell as a formatted string.
    pub fn get_value(&self, sheet: &str, row: u32, col: u32) -> Option<String> {
        use formualizer_workbook::LiteralValue;

        let value = self.workbook.get_value(sheet, row, col)?;

        Some(match value {
            LiteralValue::Empty => String::new(),
            LiteralValue::Int(i) => i.to_string(),
            LiteralValue::Number(n) => {
                // Format without unnecessary decimal places
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{:.0}", n)
                } else {
                    n.to_string()
                }
            }
            LiteralValue::Text(s) => s,
            LiteralValue::Boolean(b) => if b { "TRUE" } else { "FALSE" }.to_string(),
            LiteralValue::Date(d) => d.format("%Y-%m-%d").to_string(),
            LiteralValue::DateTime(dt) => dt.format("%Y-%m-%d %H:%M:%S").to_string(),
            LiteralValue::Time(t) => t.format("%H:%M:%S").to_string(),
            LiteralValue::Duration(dur) => {
                let secs = dur.num_seconds();
                let hours = secs / 3600;
                let mins = (secs % 3600) / 60;
                let secs = secs % 60;
                format!("{:02}:{:02}:{:02}", hours, mins, secs)
            }
            LiteralValue::Error(e) => format!("{}", e),
            LiteralValue::Array(arr) => {
                // Format array as semicolon-separated rows, comma-separated cols
                arr.iter()
                    .map(|row| {
                        row.iter()
                            .map(|v| format_literal(v))
                            .collect::<Vec<_>>()
                            .join(",")
                    })
                    .collect::<Vec<_>>()
                    .join(";")
            }
            LiteralValue::Pending => String::from("#PENDING"),
        })
    }

    /// Get list of sheet names.
    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    /// Get the dimensions (max_col, max_row) of a sheet.
    pub fn get_sheet_dimensions(&self, sheet: &str) -> Option<(u32, u32)> {
        // Get dimensions from the engine's sheet store
        let sheet_store = self.workbook.engine().sheet_store();
        let asheet = sheet_store.sheet(sheet)?;

        // Get the bounds from the arrow sheet
        let rows = asheet.nrows;
        let cols = asheet.columns.len() as u32;
        Some((cols, rows))
    }
}

/// Format a LiteralValue to string (helper for arrays).
fn format_literal(value: &formualizer_workbook::LiteralValue) -> String {
    use formualizer_workbook::LiteralValue;

    match value {
        LiteralValue::Empty => String::new(),
        LiteralValue::Int(i) => i.to_string(),
        LiteralValue::Number(n) => {
            if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{:.0}", n)
            } else {
                n.to_string()
            }
        }
        LiteralValue::Text(s) => s.clone(),
        LiteralValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
        LiteralValue::Date(d) => d.format("%Y-%m-%d").to_string(),
        LiteralValue::DateTime(dt) => dt.format("%Y-%m-%d %H:%M:%S").to_string(),
        LiteralValue::Time(t) => t.format("%H:%M:%S").to_string(),
        LiteralValue::Duration(dur) => {
            let secs = dur.num_seconds();
            format!("{}s", secs)
        }
        LiteralValue::Error(e) => format!("{}", e),
        LiteralValue::Array(_) => String::from("{...}"),
        LiteralValue::Pending => String::from("#PENDING"),
    }
}

/// Parse a cell reference like "Sheet1!A1" or "A1" and apply an override.
fn apply_override(
    workbook: &mut Workbook,
    cell_ref: &str,
    value: &str,
    sheet_names: &[String],
) -> Result<()> {
    let (sheet, col, row) = parse_cell_reference(cell_ref, sheet_names)?;

    // Parse value to appropriate type
    let literal = parse_value(value);

    workbook
        .set_value(&sheet, row, col, literal)
        .map_err(|e| Error::InvalidExcel {
            path: std::path::PathBuf::new(),
            details: format!("failed to set {}: {}", cell_ref, e),
        })?;

    Ok(())
}

/// Parse a cell reference string into (sheet, col, row).
fn parse_cell_reference(cell_ref: &str, sheet_names: &[String]) -> Result<(String, u32, u32)> {
    let (sheet, cell) = if let Some(idx) = cell_ref.find('!') {
        let sheet = &cell_ref[..idx];
        let cell = &cell_ref[idx + 1..];
        (sheet.to_string(), cell)
    } else {
        // Use first sheet if not specified
        let sheet = sheet_names
            .first()
            .ok_or_else(|| Error::SheetNotFound {
                name: "default".to_string(),
                available: String::new(),
            })?
            .clone();
        (sheet, cell_ref)
    };

    let (col, row) = parse_a1_notation(cell)?;
    Ok((sheet, col, row))
}

/// Parse A1 notation (e.g., "A1", "BC123") into (col, row).
fn parse_a1_notation(cell: &str) -> Result<(u32, u32)> {
    let cell = cell.trim().to_uppercase();

    // Find where letters end and digits begin
    let col_end = cell
        .chars()
        .position(|c| c.is_ascii_digit())
        .ok_or_else(|| Error::InvalidExcel {
            path: std::path::PathBuf::new(),
            details: format!("invalid cell reference: {}", cell),
        })?;

    let col_str = &cell[..col_end];
    let row_str = &cell[col_end..];

    // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
    let col: u32 = col_str.chars().fold(0, |acc, c| {
        acc * 26 + (c as u32 - 'A' as u32 + 1)
    });

    let row: u32 = row_str.parse().map_err(|_| Error::InvalidExcel {
        path: std::path::PathBuf::new(),
        details: format!("invalid row in cell reference: {}", cell),
    })?;

    Ok((col, row))
}

/// Parse a string value into a LiteralValue.
fn parse_value(value: &str) -> formualizer_workbook::LiteralValue {
    use formualizer_workbook::LiteralValue;

    // Try boolean
    if value.eq_ignore_ascii_case("true") {
        return LiteralValue::Boolean(true);
    }
    if value.eq_ignore_ascii_case("false") {
        return LiteralValue::Boolean(false);
    }

    // Try integer
    if let Ok(i) = value.parse::<i64>() {
        return LiteralValue::Int(i);
    }

    // Try float
    if let Ok(n) = value.parse::<f64>() {
        return LiteralValue::Number(n);
    }

    // Try date (YYYY-MM-DD)
    if let Ok(d) = NaiveDate::parse_from_str(value, "%Y-%m-%d") {
        return LiteralValue::Date(d);
    }

    // Default to text
    LiteralValue::Text(value.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_a1_notation() {
        assert_eq!(parse_a1_notation("A1").unwrap(), (1, 1));
        assert_eq!(parse_a1_notation("B2").unwrap(), (2, 2));
        assert_eq!(parse_a1_notation("Z1").unwrap(), (26, 1));
        assert_eq!(parse_a1_notation("AA1").unwrap(), (27, 1));
        assert_eq!(parse_a1_notation("AB10").unwrap(), (28, 10));
    }

    #[test]
    fn test_parse_value() {
        use formualizer_workbook::LiteralValue;

        assert!(matches!(parse_value("true"), LiteralValue::Boolean(true)));
        assert!(matches!(parse_value("FALSE"), LiteralValue::Boolean(false)));
        assert!(matches!(parse_value("42"), LiteralValue::Int(42)));
        assert!(matches!(parse_value("3.14"), LiteralValue::Number(n) if (n - 3.14).abs() < 0.001));
        assert!(matches!(parse_value("hello"), LiteralValue::Text(s) if s == "hello"));
    }
}
