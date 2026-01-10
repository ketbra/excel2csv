//! Creates a test Excel file with edge cases for formula recalculation testing.

use umya_spreadsheet::*;

fn main() {
    let mut book = new_file();

    {
        let sheet = book.get_sheet_mut(&0).unwrap();
        sheet.set_name("EdgeCases");

        // Headers
        sheet.get_cell_mut("A1").set_value("Test Case");
        sheet.get_cell_mut("B1").set_value("Input");
        sheet.get_cell_mut("C1").set_value("Formula");
        sheet.get_cell_mut("D1").set_value("Expected");

        // Test 1: Basic arithmetic
        sheet.get_cell_mut("A2").set_value("Basic math");
        sheet.get_cell_mut("B2").set_value_number(100.0);
        sheet.get_cell_mut("C2").set_formula("=B2*2");
        sheet.get_cell_mut("D2").set_value("200");

        // Test 2: Division (potential div/0)
        sheet.get_cell_mut("A3").set_value("Division");
        sheet.get_cell_mut("B3").set_value_number(50.0);
        sheet.get_cell_mut("C3").set_formula("=B3/10");
        sheet.get_cell_mut("D3").set_value("5");

        // Test 3: Division by zero
        sheet.get_cell_mut("A4").set_value("Div by zero");
        sheet.get_cell_mut("B4").set_value_number(0.0);
        sheet.get_cell_mut("C4").set_formula("=100/B4");
        sheet.get_cell_mut("D4").set_value("#DIV/0!");

        // Test 4: Negative numbers
        sheet.get_cell_mut("A5").set_value("Negative");
        sheet.get_cell_mut("B5").set_value_number(-25.0);
        sheet.get_cell_mut("C5").set_formula("=B5*2");
        sheet.get_cell_mut("D5").set_value("-50");

        // Test 5: String concatenation
        sheet.get_cell_mut("A6").set_value("Concat");
        sheet.get_cell_mut("B6").set_value("Hello");
        sheet.get_cell_mut("C6").set_formula("=B6&\" World\"");
        sheet.get_cell_mut("D6").set_value("Hello World");

        // Test 6: Nested IF
        sheet.get_cell_mut("A7").set_value("Nested IF");
        sheet.get_cell_mut("B7").set_value_number(75.0);
        sheet.get_cell_mut("C7").set_formula("=IF(B7>=90,\"A\",IF(B7>=80,\"B\",IF(B7>=70,\"C\",\"F\")))");
        sheet.get_cell_mut("D7").set_value("C");

        // Test 7: Empty cell reference
        sheet.get_cell_mut("A8").set_value("Empty ref");
        // B8 is intentionally empty
        sheet.get_cell_mut("C8").set_formula("=B8+10");
        sheet.get_cell_mut("D8").set_value("10");

        // Test 8: Boolean
        sheet.get_cell_mut("A9").set_value("Boolean");
        sheet.get_cell_mut("B9").set_value_number(10.0);
        sheet.get_cell_mut("C9").set_formula("=B9>5");
        sheet.get_cell_mut("D9").set_value("TRUE");

        // Test 9: Circular reference (should error)
        sheet.get_cell_mut("A10").set_value("Circular");
        sheet.get_cell_mut("C10").set_formula("=C10+1");
        sheet.get_cell_mut("D10").set_value("Error");

        // Test 10: Cross-row reference
        sheet.get_cell_mut("A11").set_value("Sum range");
        sheet.get_cell_mut("C11").set_formula("=SUM(B2:B7)");
        sheet.get_cell_mut("D11").set_value("200"); // 100+50+0+(-25)+0+75
    }

    // Save the workbook
    let path = "edge_case_test.xlsx";
    writer::xlsx::write(&book, path).expect("Failed to write edge_case_test.xlsx");
    println!("Created edge_case_test.xlsx successfully!");
    println!();
    println!("Test cases:");
    println!("  1. Basic math: 100*2 = 200");
    println!("  2. Division: 50/10 = 5");
    println!("  3. Div by zero: 100/0 = #DIV/0!");
    println!("  4. Negative: -25*2 = -50");
    println!("  5. Concat: \"Hello\" & \" World\" = \"Hello World\"");
    println!("  6. Nested IF: 75 -> \"C\"");
    println!("  7. Empty ref: (empty)+10 = 10");
    println!("  8. Boolean: 10>5 = TRUE");
    println!("  9. Circular: =C10+1 -> Error");
    println!(" 10. Sum range: SUM(B2:B7) = 200");
}
