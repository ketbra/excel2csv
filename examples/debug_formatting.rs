// Diagnostic script to understand formatting behavior
use std::path::Path;

fn main() {
    let path = Path::new("sample.xlsx");
    let workbook = umya_spreadsheet::reader::xlsx::read(path).expect("Failed to read file");

    // Get the first sheet
    let sheet = workbook.get_sheet(&0).expect("No sheet found");
    println!("Sheet name: {}", sheet.get_name());

    // Look at cells that might have the {1} format
    // Based on screenshot, looks like column B has these values
    println!("\n=== Examining cells with potential custom formatting ===\n");

    // Check a range of cells in column B (column index 2)
    for row in 1..=30 {
        for col in 1..=5 {
            if let Some(cell) = sheet.get_cell((col, row)) {
                let raw_value = cell.get_value();
                let formatted_value = sheet.get_formatted_value((col, row));

                // Get the style/format info
                let style = cell.get_style();
                let format_code = style
                    .get_number_format()
                    .map(|nf| nf.get_format_code())
                    .unwrap_or("General");

                // Only print if there's interesting data
                if !raw_value.is_empty() && (format_code != "General" || raw_value != formatted_value) {
                    println!(
                        "Cell ({}, {}): raw='{}', formatted='{}', format_code='{}'",
                        col, row, raw_value, formatted_value, format_code
                    );
                }
            }
        }
    }

    // Also specifically look for cells with custom format containing braces
    println!("\n=== Searching for cells with brace formatting ===\n");
    let cells = sheet.get_cell_collection();
    for cell in cells {
        let coord = cell.get_coordinate();
        let style = cell.get_style();
        let format_code = style
            .get_number_format()
            .map(|nf| nf.get_format_code())
            .unwrap_or("General");

        if format_code.contains("{") || format_code.contains("}") || format_code.contains("\\") {
            let raw = cell.get_value();
            let formatted = sheet.get_formatted_value(coord.get_coordinate());
            println!(
                "Cell {}: raw='{}', formatted='{}', format_code='{}'",
                coord.get_coordinate(), raw, formatted, format_code
            );
        }
    }
}
