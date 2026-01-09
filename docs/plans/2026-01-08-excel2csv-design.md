# excel2csv Design

A Rust CLI tool for converting Excel files (.xls and .xlsx) to CSV with formatted output.

## Requirements

- Support XLS (BIFF8) and XLSX formats
- Preserve formatted cell values (e.g., `$3,123.45` not `3123.45`)
- Preserve date formatting from spreadsheet
- Multi-sheet support with option to select specific sheet

## CLI Interface

```
excel2csv [OPTIONS] <INPUT_FILE>

Arguments:
  <INPUT_FILE>    Excel file to convert (.xls or .xlsx)

Options:
  -o, --output <PATH>      Output file or directory (default: stdout)
  -s, --sheet <NAME|INDEX> Convert specific sheet (default: all sheets)
  -f, --format <FORMAT>    Output format: csv, tsv, european (default: csv)
  -e, --empty <STRING>     Value for empty cells (default: "")
  -v, --verbose            Print detailed progress to stderr
  -h, --help               Print help
  -V, --version            Print version
```

### Output Behavior

- Single sheet + stdout: CSV to stdout
- Single sheet + `-o file.csv`: CSV to file
- Multiple sheets + stdout: Error, must specify `-o directory/`
- Multiple sheets + `-o dir/`: `dir/Sheet1.csv`, `dir/Sheet2.csv`, etc.

### Format Presets

- `csv`: Comma delimiter (default)
- `tsv`: Tab delimiter
- `european`: Semicolon delimiter

## Architecture

```
┌─────────────┐     ┌─────────────────────────────┐     ┌─────────────┐
│  .xls file  │────>│  xlrd::open()               │────>│             │
└─────────────┘     │  -> umya_spreadsheet::      │     │             │
                    │     Spreadsheet             │     │             │
                    └─────────────────────────────┘     │  Spreadsheet │
                                                        │  (unified)   │
┌─────────────┐     ┌─────────────────────────────┐     │             │
│ .xlsx file  │────>│  umya_spreadsheet::reader   │────>│             │
└─────────────┘     │  ::xlsx::read()             │     │             │
                    └─────────────────────────────┘     └──────┬──────┘
                                                               │
                                                               v
                                                        ┌─────────────┐
                                                        │ CSV Writer  │
                                                        │ (per sheet) │
                                                        └─────────────┘
```

- File type detected by extension (`.xls` vs `.xlsx`/`.xlsm`)
- Both paths produce the same `Spreadsheet` type
- Formatted cell values via `worksheet.get_formatted_value()`
- CSV written using the `csv` crate

## Formatted Value Extraction

Uses `umya_spreadsheet::Worksheet::get_formatted_value(coordinate)` which returns the display string with formatting applied.

Per-row iteration:
1. Determine sheet dimensions (max row/column with data)
2. For each row, iterate columns
3. Call `get_formatted_value()` for each cell
4. Empty cells return empty string (or `--empty` value if specified)
5. Write row to CSV writer

Edge cases:
- Cells with errors (`#REF!`, `#DIV/0!`) - output the error string
- Merged cells - output value only in top-left cell, empty for others
- Large sheets - stream row-by-row

## Error Handling

### Exit Codes

- `0` - Success
- `1` - Input file not found or unreadable
- `2` - Invalid/corrupt Excel file
- `3` - Invalid arguments (e.g., `--sheet` doesn't exist, multi-sheet to stdout without `-o`)

### Error Messages

```
error: file not found: input.xlsx
error: invalid Excel file: input.xls (unsupported BIFF version)
error: sheet "Sales" not found (available: Sheet1, Sheet2, Summary)
error: multiple sheets require -o <directory>
```

### Verbose Output

With `--verbose`:
```
reading: input.xlsx
format: xlsx
sheets: Sheet1, Sheet2, Summary
processing: Sheet1 (1500 rows, 12 columns)
output: ./output/Sheet1.csv
```

## Dependencies

```toml
[dependencies]
umya-spreadsheet = "2"
xlrd = { git = "https://github.com/PleaseDont/xlrd" }
csv = "1"
clap = { version = "4", features = ["derive"] }
```

## Project Structure

```
excel2csv/
├── Cargo.toml
├── src/
│   ├── main.rs         # CLI parsing, orchestration
│   ├── reader.rs       # open_workbook() - dispatches based on extension
│   ├── writer.rs       # CSV writing, format presets, empty cell handling
│   └── error.rs        # Error types, exit codes
```

~300-400 lines estimated.
