use crate::error::Result;
use crate::evaluator::EvaluatedWorkbook;
use crate::OutputFormat;
use csv::WriterBuilder;
use ssfmt::{FormatOptions, NumberFormat};
use std::io::Write;
use umya_spreadsheet::structs::Worksheet;

pub struct CsvConfig {
    pub format: OutputFormat,
    pub empty_value: String,
}

pub fn write_sheet<W: Write>(
    sheet: &Worksheet,
    writer: W,
    config: &CsvConfig,
) -> Result<()> {
    let mut csv_writer = WriterBuilder::new()
        .delimiter(config.format.delimiter())
        .from_writer(writer);

    let (max_col, max_row) = sheet.get_highest_column_and_row();
    let opts = FormatOptions::default();

    for row in 1..=max_row {
        let mut record: Vec<String> = Vec::with_capacity(max_col as usize);

        for col in 1..=max_col {
            let value = get_cell_formatted_value(sheet, col, row, &opts);
            if value.is_empty() {
                record.push(config.empty_value.clone());
            } else {
                record.push(value);
            }
        }

        csv_writer.write_record(&record)?;
    }

    csv_writer.flush()?;
    Ok(())
}

/// Write an evaluated sheet to CSV output.
pub fn write_evaluated_sheet<W: Write>(
    workbook: &EvaluatedWorkbook,
    sheet_name: &str,
    writer: W,
    config: &CsvConfig,
) -> Result<()> {
    let mut csv_writer = WriterBuilder::new()
        .delimiter(config.format.delimiter())
        .from_writer(writer);

    let (max_col, max_row) = workbook
        .get_sheet_dimensions(sheet_name)
        .unwrap_or((0, 0));

    for row in 1..=max_row {
        let mut record: Vec<String> = Vec::with_capacity(max_col as usize);

        for col in 1..=max_col {
            let value = workbook
                .get_value(sheet_name, row, col)
                .unwrap_or_default();

            if value.is_empty() {
                record.push(config.empty_value.clone());
            } else {
                record.push(value);
            }
        }

        csv_writer.write_record(&record)?;
    }

    csv_writer.flush()?;
    Ok(())
}

/// Get a cell's formatted value using ssfmt for proper Excel format code support.
fn get_cell_formatted_value(sheet: &Worksheet, col: u32, row: u32, opts: &FormatOptions) -> String {
    let Some(cell) = sheet.get_cell((col, row)) else {
        return String::new();
    };

    let raw_value = cell.get_value();
    if raw_value.is_empty() {
        return String::new();
    }

    // Get the format code
    let format_code = cell
        .get_style()
        .get_number_format()
        .map(|nf| nf.get_format_code())
        .unwrap_or("General");

    // Parse the format code
    let fmt = match NumberFormat::parse(format_code) {
        Ok(f) => f,
        Err(_) => return raw_value.to_string(),
    };

    // Try to parse as number (includes dates which are serial numbers in Excel)
    if let Ok(num) = raw_value.parse::<f64>() {
        return fmt.format(num, opts);
    }

    // Handle as text value
    fmt.format_text(&raw_value, opts)
}

