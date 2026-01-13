//! Inspect cells in an Excel file to see raw values, formulas, and formatting
use std::env;
use std::path::Path;

fn main() {
    let args: Vec<String> = env::args().collect();
    let path = Path::new(args.get(1).map(|s| s.as_str()).unwrap_or("formula_test.xlsx"));

    let workbook = umya_spreadsheet::reader::xlsx::read(path).expect("Failed to read file");

    let sheet = workbook.get_sheet(&0).expect("No sheet found");
    println!("Sheet name: {}", sheet.get_name());
    println!();
    println!("Format: (col,row) raw='value' formula='formula' format='code'");
    println!();

    for row in 1..=6 {
        for col in 1..=5 {
            if let Some(cell) = sheet.get_cell((col, row)) {
                let raw_value = cell.get_value();
                let formula = cell.get_formula();
                let style = cell.get_style();
                let format_code = style
                    .get_number_format()
                    .map(|nf| nf.get_format_code().to_string())
                    .unwrap_or_else(|| "(none)".to_string());

                println!(
                    "({},{}) raw='{}' formula='{}' format='{}'",
                    col, row, raw_value, formula, format_code
                );
            }
        }
    }
}
