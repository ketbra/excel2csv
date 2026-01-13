//! Creates a test Excel file with formulas for testing recalculation.

use umya_spreadsheet::*;

fn main() {
    let mut book = new_file();

    // Get Sheet1 (created by default) and populate it
    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("Calculations");

        // Headers
        sheet.get_cell_mut("A1").set_value("Value A");
        sheet.get_cell_mut("B1").set_value("Value B");
        sheet.get_cell_mut("C1").set_value("Sum");
        sheet.get_cell_mut("D1").set_value("Product");
        sheet.get_cell_mut("E1").set_value("Average");

        // Row 2: Values and formulas
        sheet.get_cell_mut("A2").set_value_number(10.0);
        sheet.get_cell_mut("B2").set_value_number(20.0);
        sheet.get_cell_mut("C2").set_formula("=A2+B2");
        sheet.get_cell_mut("D2").set_formula("=A2*B2");
        sheet.get_cell_mut("E2").set_formula("=AVERAGE(A2,B2)");

        // Row 3: More values and formulas
        sheet.get_cell_mut("A3").set_value_number(5.0);
        sheet.get_cell_mut("B3").set_value_number(15.0);
        sheet.get_cell_mut("C3").set_formula("=A3+B3");
        sheet.get_cell_mut("D3").set_formula("=A3*B3");
        sheet.get_cell_mut("E3").set_formula("=AVERAGE(A3,B3)");

        // Row 4: Totals
        sheet.get_cell_mut("A4").set_value("Totals:");
        sheet.get_cell_mut("C4").set_formula("=SUM(C2:C3)");
        sheet.get_cell_mut("D4").set_formula("=SUM(D2:D3)");
        sheet.get_cell_mut("E4").set_formula("=AVERAGE(E2:E3)");

        // Row 6: Nested formulas
        sheet.get_cell_mut("A6").set_value("Nested:");
        sheet.get_cell_mut("B6").set_formula("=C4+D4");
        sheet.get_cell_mut("C6").set_formula("=IF(B6>100,\"High\",\"Low\")");
    }

    // Save the workbook
    let path = "formula_test.xlsx";
    writer::xlsx::write(&book, path).expect("Failed to write formula_test.xlsx");
    println!("Created formula_test.xlsx successfully!");
    println!();
    println!("Expected values (when A2=10, B2=20, A3=5, B3=15):");
    println!("  C2 (A2+B2) = 30");
    println!("  D2 (A2*B2) = 200");
    println!("  E2 (AVERAGE) = 15");
    println!("  C3 (A3+B3) = 20");
    println!("  D3 (A3*B3) = 75");
    println!("  E3 (AVERAGE) = 10");
    println!("  C4 (SUM C2:C3) = 50");
    println!("  D4 (SUM D2:D3) = 275");
    println!("  E4 (AVERAGE E2:E3) = 12.5");
    println!("  B6 (C4+D4) = 325");
    println!("  C6 (IF) = High");
}
