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

    // Try our custom format handling first for cases umya-spreadsheet doesn't handle well
    if let Some(formatted) = apply_format_with_sections(&raw_value, format_code) {
        return formatted;
    }

    // Fall back to library's formatting
    sheet.get_formatted_value((col, row))
}

/// Apply format handling with support for multi-section formats and literal text.
/// Excel format codes can have up to 4 sections separated by semicolons:
///   1 section:  [POSITIVE/NEGATIVE/ZERO/TEXT]
///   2 sections: [POSITIVE/ZERO] [NEGATIVE]
///   3 sections: [POSITIVE] [NEGATIVE] [ZERO]
///   4 sections: [POSITIVE] [NEGATIVE] [ZERO] [TEXT]
fn apply_format_with_sections(raw_value: &str, format_code: &str) -> Option<String> {
    // Parse the number
    let num: f64 = raw_value.parse().ok()?;

    // Split format into sections (but not within quoted strings)
    let sections = split_format_sections(format_code);

    // Select the appropriate section based on value
    let section = select_format_section(&sections, num);

    // Check if this section needs custom handling
    if needs_custom_handling(section) {
        return apply_section_format(num, section);
    }

    // Let the library handle standard formats
    None
}

/// Split format code into sections, respecting quoted strings
fn split_format_sections(format_code: &str) -> Vec<&str> {
    let mut sections = Vec::new();
    let mut start = 0;
    let mut in_quotes = false;

    for (i, c) in format_code.char_indices() {
        match c {
            '"' => in_quotes = !in_quotes,
            ';' if !in_quotes => {
                sections.push(&format_code[start..i]);
                start = i + 1;
            }
            _ => {}
        }
    }
    sections.push(&format_code[start..]);
    sections
}

/// Select the appropriate format section based on the numeric value
fn select_format_section<'a>(sections: &[&'a str], value: f64) -> &'a str {
    match sections.len() {
        1 => sections[0],
        2 => {
            if value >= 0.0 {
                sections[0]
            } else {
                sections[1]
            }
        }
        3 | 4 => {
            if value > 0.0 {
                sections[0]
            } else if value < 0.0 {
                sections[1]
            } else {
                sections[2] // Zero
            }
        }
        _ => sections[0],
    }
}

/// Check if a format section needs custom handling
fn needs_custom_handling(section: &str) -> bool {
    // Handle literal characters that umya-spreadsheet doesn't process well
    let has_custom_literals = section.contains("\\{") || section.contains("\\}");

    // Handle zero sections with quoted literals (like "-" for displaying hyphen)
    let has_quoted_literal_only = section.contains('"')
        && !section.contains('#')
        && !section.contains('0');

    // Handle sections that are purely literal text
    let is_literal_section = has_quoted_literal_only
        || (has_custom_literals && !section.contains('#') && !section.contains('0'));

    has_custom_literals || has_quoted_literal_only || is_literal_section
}

/// Apply formatting for a single section
fn apply_section_format(num: f64, section: &str) -> Option<String> {
    let mut result = String::new();
    let mut number_format = String::new();
    let mut prefix = String::new();
    let mut suffix = String::new();
    let mut in_number_section = false;
    let mut has_number_placeholder = false;

    let chars: Vec<char> = section.chars().collect();
    let mut i = 0;

    while i < chars.len() {
        let c = chars[i];

        if c == '\\' && i + 1 < chars.len() {
            // Escaped literal character
            let literal = chars[i + 1];
            if in_number_section || has_number_placeholder {
                suffix.push(literal);
            } else {
                prefix.push(literal);
            }
            i += 2;
            continue;
        }

        if c == '"' {
            // Quoted string - extract content
            i += 1;
            while i < chars.len() && chars[i] != '"' {
                if in_number_section || has_number_placeholder {
                    suffix.push(chars[i]);
                } else {
                    prefix.push(chars[i]);
                }
                i += 1;
            }
            i += 1;
            continue;
        }

        if c == '#' || c == '0' {
            in_number_section = true;
            has_number_placeholder = true;
            number_format.push(c);
        } else if c == '.' || c == ',' {
            if in_number_section {
                number_format.push(c);
            }
        } else if c == '?' {
            // Optional digit placeholder - treat like #
            if in_number_section || has_number_placeholder {
                // In suffix area, ignore
            } else {
                // Could be alignment placeholder
            }
        } else if c == '_' || c == '*' {
            // Skip spacing/fill characters and the next char
            i += 2;
            continue;
        } else if c == '[' {
            // Skip bracketed content like [$-409]
            while i < chars.len() && chars[i] != ']' {
                i += 1;
            }
        } else if !in_number_section && !has_number_placeholder {
            // Before number format - add to prefix if significant
            if !c.is_whitespace() && c != '(' && c != ')' {
                prefix.push(c);
            }
        } else {
            // After number format
            if !c.is_whitespace() {
                suffix.push(c);
                in_number_section = false;
            }
        }

        i += 1;
    }

    // If no number placeholders, this is a literal-only section
    if !has_number_placeholder {
        // Return just the literals (prefix + suffix)
        let literal = format!("{}{}", prefix, suffix);
        if !literal.is_empty() {
            return Some(literal);
        }
        return None;
    }

    // Format the number
    let use_thousands = number_format.contains(',');
    let formatted_num = if number_format.contains('.') {
        let decimal_places = number_format
            .split('.')
            .nth(1)
            .map(|s| s.chars().filter(|c| *c == '0' || *c == '#').count())
            .unwrap_or(0);
        let abs_num = num.abs();
        if use_thousands {
            format_with_thousands(abs_num, decimal_places)
        } else {
            format!("{:.prec$}", abs_num, prec = decimal_places)
        }
    } else {
        let abs_num = num.abs() as i64;
        if use_thousands {
            format_int_with_thousands(abs_num)
        } else {
            format!("{}", abs_num)
        }
    };

    result.push_str(&prefix);
    result.push_str(&formatted_num);
    result.push_str(&suffix);

    Some(result)
}

/// Format a number with thousands separators
fn format_with_thousands(num: f64, decimal_places: usize) -> String {
    let formatted = format!("{:.prec$}", num, prec = decimal_places);
    let parts: Vec<&str> = formatted.split('.').collect();
    let int_part = parts[0];
    let dec_part = parts.get(1);

    let int_with_commas = format_int_with_thousands(int_part.parse().unwrap_or(0));

    match dec_part {
        Some(d) => format!("{}.{}", int_with_commas, d),
        None => int_with_commas,
    }
}

/// Format an integer with thousands separators
fn format_int_with_thousands(num: i64) -> String {
    let s = num.to_string();
    let mut result = String::new();
    for (i, c) in s.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            result.push(',');
        }
        result.push(c);
    }
    result.chars().rev().collect()
}
