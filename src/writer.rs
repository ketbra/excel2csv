use crate::error::Result;
use crate::OutputFormat;
use csv::WriterBuilder;
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

    for row in 1..=max_row {
        let mut record: Vec<String> = Vec::with_capacity(max_col as usize);

        for col in 1..=max_col {
            let value = get_cell_formatted_value(sheet, col, row);
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

/// Get a cell's formatted value, with custom handling for format codes
/// that umya-spreadsheet doesn't fully support (like literal text).
fn get_cell_formatted_value(sheet: &Worksheet, col: u32, row: u32) -> String {
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

    // Check if format contains escaped literal characters that umya-spreadsheet
    // doesn't handle properly (like \{ or \})
    if format_code.contains('\\') && !is_standard_format(format_code) {
        if let Some(formatted) = apply_custom_format(&raw_value, format_code) {
            return formatted;
        }
    }

    // Fall back to library's formatting
    sheet.get_formatted_value((col, row))
}

/// Check if the format code is a standard format that umya-spreadsheet handles well
fn is_standard_format(format_code: &str) -> bool {
    // Standard accounting/currency formats with escaped parens are handled OK
    // But custom literal formats like \{###\} are not
    let dominated_by_parens = format_code.matches("\\(").count() > 0
        && format_code.matches("\\)").count() > 0
        && !format_code.contains("\\{")
        && !format_code.contains("\\}");

    dominated_by_parens || format_code == "General"
}

/// Apply a custom format code that contains literal text
fn apply_custom_format(raw_value: &str, format_code: &str) -> Option<String> {
    // Parse the number
    let num: f64 = raw_value.parse().ok()?;

    // Handle format codes with escaped literals like \{###\}
    // Strategy: extract prefix, number format, suffix from the format code

    let mut result = String::new();
    let mut in_number_section = false;
    let mut number_format = String::new();
    let mut prefix = String::new();
    let mut suffix = String::new();

    let chars: Vec<char> = format_code.chars().collect();
    let mut i = 0;

    while i < chars.len() {
        let c = chars[i];

        if c == '\\' && i + 1 < chars.len() {
            // Escaped literal character
            let literal = chars[i + 1];
            if in_number_section {
                // We've seen number chars, so this is suffix
                suffix.push(literal);
                in_number_section = false; // switch to suffix mode
            } else if number_format.is_empty() {
                prefix.push(literal);
            } else {
                suffix.push(literal);
            }
            i += 2;
            continue;
        }

        if c == '#' || c == '0' || c == '.' || c == ',' {
            in_number_section = true;
            number_format.push(c);
        } else if c == '"' {
            // Quoted string - find the closing quote
            i += 1;
            while i < chars.len() && chars[i] != '"' {
                if in_number_section || !number_format.is_empty() {
                    suffix.push(chars[i]);
                } else {
                    prefix.push(chars[i]);
                }
                i += 1;
            }
        } else if !in_number_section && number_format.is_empty() {
            // Before number format
            if c != '_' && c != '*' && c != ' ' {
                prefix.push(c);
            }
        } else {
            // After number format
            if c != '_' && c != '*' && c != ' ' {
                suffix.push(c);
            }
        }

        i += 1;
    }

    // Format the number
    let formatted_num = if number_format.contains('.') {
        // Count decimal places
        let decimal_places = number_format
            .split('.')
            .nth(1)
            .map(|s| s.chars().filter(|c| *c == '0' || *c == '#').count())
            .unwrap_or(0);
        format!("{:.prec$}", num, prec = decimal_places)
    } else {
        // Integer format
        format!("{}", num as i64)
    };

    result.push_str(&prefix);
    result.push_str(&formatted_num);
    result.push_str(&suffix);

    Some(result)
}
