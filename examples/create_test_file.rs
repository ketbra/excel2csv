use umya_spreadsheet::*;

fn main() {
    let mut book = new_file();

    // Get Sheet1 (created by default) and populate it
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Sales Data");

        // Headers
        sheet.get_cell_mut("A1").set_value("Product");
        sheet.get_cell_mut("B1").set_value("Price");
        sheet.get_cell_mut("C1").set_value("Quantity");
        sheet.get_cell_mut("D1").set_value("Date");
        sheet.get_cell_mut("E1").set_value("Notes");

        // Row 2: Various data types
        sheet.get_cell_mut("A2").set_value("Widget A");
        sheet.get_cell_mut("B2").set_value_number(19.99);
        sheet.get_cell_mut("C2").set_value_number(100);
        sheet.get_cell_mut("D2").set_value("2024-01-15");
        sheet.get_cell_mut("E2").set_value("Best seller");

        // Row 3: More data with an empty cell
        sheet.get_cell_mut("A3").set_value("Gadget B");
        sheet.get_cell_mut("B3").set_value_number(49.50);
        sheet.get_cell_mut("C3").set_value_number(25);
        sheet.get_cell_mut("D3").set_value("2024-02-20");
        // E3 is intentionally empty

        // Row 4: Contains special characters
        sheet.get_cell_mut("A4").set_value("Item \"C\"");
        sheet.get_cell_mut("B4").set_value_number(9.99);
        sheet.get_cell_mut("C4").set_value_number(500);
        sheet.get_cell_mut("D4").set_value("2024-03-10");
        sheet.get_cell_mut("E4").set_value("Has, comma");

        // Row 5: Large numbers
        sheet.get_cell_mut("A5").set_value("Bulk Order");
        sheet.get_cell_mut("B5").set_value_number(1234567.89);
        sheet.get_cell_mut("C5").set_value_number(10000);
        sheet.get_cell_mut("D5").set_value("2024-04-01");
        sheet.get_cell_mut("E5").set_value("Wholesale");
    }

    // Create Sheet2
    book.new_sheet("Inventory");
    {
        let sheet = book.get_sheet_by_name_mut("Inventory").unwrap();

        // Headers
        sheet.get_cell_mut("A1").set_value("SKU");
        sheet.get_cell_mut("B1").set_value("In Stock");
        sheet.get_cell_mut("C1").set_value("Reorder Level");

        // Data rows
        sheet.get_cell_mut("A2").set_value("SKU-001");
        sheet.get_cell_mut("B2").set_value_number(150);
        sheet.get_cell_mut("C2").set_value_number(50);

        sheet.get_cell_mut("A3").set_value("SKU-002");
        sheet.get_cell_mut("B3").set_value_number(75);
        sheet.get_cell_mut("C3").set_value_number(100);

        sheet.get_cell_mut("A4").set_value("SKU-003");
        sheet.get_cell_mut("B4").set_value_number(0);
        sheet.get_cell_mut("C4").set_value_number(25);
    }

    // Save the workbook
    let path = "test.xlsx";
    writer::xlsx::write(&book, path).expect("Failed to write test.xlsx");
    println!("Created test.xlsx successfully!");
    println!("Sheets: Sales Data, Inventory");
}
