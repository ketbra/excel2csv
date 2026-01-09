# excel2csv Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a CLI tool that converts Excel files (.xls, .xlsx) to CSV with formatted cell values preserved.

**Architecture:** Detect file type by extension, use xlrd for .xls and umya-spreadsheet for .xlsx, both producing a unified Spreadsheet type. Iterate cells using get_formatted_value() and write to CSV.

**Tech Stack:** Rust, clap (CLI), umya-spreadsheet (XLSX + data model), xlrd (XLS), csv (output)

---

## Task 1: Project Setup

**Files:**
- Create: `Cargo.toml`
- Create: `src/main.rs`
- Create: `src/error.rs`

**Step 1: Create Cargo.toml**

```toml
[package]
name = "excel2csv"
version = "0.1.0"
edition = "2021"
description = "Convert Excel files to CSV with formatted output"

[dependencies]
umya-spreadsheet = "2"
xlrd = { git = "https://github.com/PleaseDont/xlrd" }
csv = "1"
clap = { version = "4", features = ["derive"] }
thiserror = "1"
```

**Step 2: Create src/error.rs**

```rust
use std::path::PathBuf;
use thiserror::Error;

#[derive(Error, Debug)]
pub enum Error {
    #[error("file not found: {0}")]
    FileNotFound(PathBuf),

    #[error("invalid Excel file: {path} ({details})")]
    InvalidExcel { path: PathBuf, details: String },

    #[error("sheet \"{name}\" not found (available: {available})")]
    SheetNotFound { name: String, available: String },

    #[error("sheet index {index} out of range (have {count} sheets)")]
    SheetIndexOutOfRange { index: usize, count: usize },

    #[error("multiple sheets require -o <directory>")]
    MultipleSheetNoOutput,

    #[error("unsupported file format: {0}")]
    UnsupportedFormat(String),

    #[error("failed to write CSV: {0}")]
    CsvWrite(#[from] csv::Error),

    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
}

impl Error {
    pub fn exit_code(&self) -> i32 {
        match self {
            Error::FileNotFound(_) => 1,
            Error::InvalidExcel { .. } => 2,
            Error::SheetNotFound { .. } => 3,
            Error::SheetIndexOutOfRange { .. } => 3,
            Error::MultipleSheetNoOutput => 3,
            Error::UnsupportedFormat(_) => 3,
            Error::CsvWrite(_) => 4,
            Error::Io(_) => 1,
        }
    }
}

pub type Result<T> = std::result::Result<T, Error>;
```

**Step 3: Create minimal src/main.rs**

```rust
mod error;

fn main() {
    println!("excel2csv");
}
```

**Step 4: Verify it compiles**

Run: `cargo build`
Expected: Compiles successfully (dependencies download)

**Step 5: Commit**

```bash
git add Cargo.toml src/
git commit -m "feat: initial project setup with dependencies"
```

---

## Task 2: CLI Argument Parsing

**Files:**
- Modify: `src/main.rs`

**Step 1: Define CLI struct with clap**

Replace `src/main.rs` with:

```rust
mod error;

use clap::{Parser, ValueEnum};
use std::path::PathBuf;

#[derive(Debug, Clone, Copy, ValueEnum, Default)]
pub enum OutputFormat {
    #[default]
    Csv,
    Tsv,
    European,
}

impl OutputFormat {
    pub fn delimiter(&self) -> u8 {
        match self {
            OutputFormat::Csv => b',',
            OutputFormat::Tsv => b'\t',
            OutputFormat::European => b';',
        }
    }
}

#[derive(Parser, Debug)]
#[command(name = "excel2csv")]
#[command(version, about = "Convert Excel files to CSV with formatted output")]
pub struct Args {
    /// Excel file to convert (.xls or .xlsx)
    pub input: PathBuf,

    /// Output file or directory (default: stdout)
    #[arg(short, long)]
    pub output: Option<PathBuf>,

    /// Convert specific sheet by name or index (default: all sheets)
    #[arg(short, long)]
    pub sheet: Option<String>,

    /// Output format
    #[arg(short, long, value_enum, default_value_t = OutputFormat::Csv)]
    pub format: OutputFormat,

    /// Value for empty cells
    #[arg(short, long, default_value = "")]
    pub empty: String,

    /// Print detailed progress to stderr
    #[arg(short, long)]
    pub verbose: bool,
}

fn main() {
    let args = Args::parse();
    if args.verbose {
        eprintln!("input: {:?}", args.input);
        eprintln!("format: {:?}", args.format);
    }
}
```

**Step 2: Verify help output**

Run: `cargo run -- --help`
Expected:
```
Convert Excel files to CSV with formatted output

Usage: excel2csv [OPTIONS] <INPUT>

Arguments:
  <INPUT>  Excel file to convert (.xls or .xlsx)

Options:
  -o, --output <OUTPUT>  Output file or directory (default: stdout)
  -s, --sheet <SHEET>    Convert specific sheet by name or index (default: all sheets)
  -f, --format <FORMAT>  Output format [default: csv] [possible values: csv, tsv, european]
  -e, --empty <EMPTY>    Value for empty cells [default: ]
  -v, --verbose          Print detailed progress to stderr
  -h, --help             Print help
  -V, --version          Print version
```

**Step 3: Commit**

```bash
git add src/main.rs
git commit -m "feat: add CLI argument parsing with clap"
```

---

## Task 3: Reader Module - XLSX Support

**Files:**
- Create: `src/reader.rs`
- Modify: `src/main.rs`

**Step 1: Create src/reader.rs with XLSX support**

```rust
use crate::error::{Error, Result};
use std::path::Path;
use umya_spreadsheet::Spreadsheet;

pub fn open_workbook(path: &Path) -> Result<Spreadsheet> {
    if !path.exists() {
        return Err(Error::FileNotFound(path.to_path_buf()));
    }

    let extension = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_default();

    match extension.as_str() {
        "xlsx" | "xlsm" => open_xlsx(path),
        "xls" => open_xls(path),
        _ => Err(Error::UnsupportedFormat(extension)),
    }
}

fn open_xlsx(path: &Path) -> Result<Spreadsheet> {
    umya_spreadsheet::reader::xlsx::read(path).map_err(|e| Error::InvalidExcel {
        path: path.to_path_buf(),
        details: e.to_string(),
    })
}

fn open_xls(path: &Path) -> Result<Spreadsheet> {
    xlrd::open(path).map_err(|e| Error::InvalidExcel {
        path: path.to_path_buf(),
        details: e.to_string(),
    })
}
```

**Step 2: Update src/main.rs to use reader**

Add after `mod error;`:

```rust
mod reader;
```

Update main function:

```rust
fn main() {
    let args = Args::parse();

    if let Err(e) = run(args) {
        eprintln!("error: {}", e);
        std::process::exit(e.exit_code());
    }
}

fn run(args: Args) -> error::Result<()> {
    if args.verbose {
        eprintln!("reading: {:?}", args.input);
    }

    let workbook = reader::open_workbook(&args.input)?;

    if args.verbose {
        eprintln!("opened successfully");
    }

    Ok(())
}
```

**Step 3: Verify it compiles**

Run: `cargo build`
Expected: Compiles successfully

**Step 4: Commit**

```bash
git add src/reader.rs src/main.rs
git commit -m "feat: add reader module for XLSX and XLS files"
```

---

## Task 4: Sheet Enumeration Helpers

**Files:**
- Modify: `src/reader.rs`

**Step 1: Add helper functions for sheet access**

Add to `src/reader.rs`:

```rust
/// Get list of sheet names from workbook
pub fn get_sheet_names(workbook: &Spreadsheet) -> Vec<String> {
    let mut names = Vec::new();
    let mut index = 0u32;
    while let Some(sheet) = workbook.get_sheet(&index) {
        names.push(sheet.get_name().to_string());
        index += 1;
    }
    names
}

/// Get sheet count
pub fn get_sheet_count(workbook: &Spreadsheet) -> usize {
    let mut count = 0;
    while workbook.get_sheet(&(count as u32)).is_some() {
        count += 1;
    }
    count
}

/// Parse sheet selector (name or index)
pub fn parse_sheet_selector(selector: &str) -> SheetSelector {
    if let Ok(index) = selector.parse::<usize>() {
        SheetSelector::Index(index)
    } else {
        SheetSelector::Name(selector.to_string())
    }
}

#[derive(Debug, Clone)]
pub enum SheetSelector {
    Index(usize),
    Name(String),
}
```

**Step 2: Verify it compiles**

Run: `cargo build`
Expected: Compiles successfully

**Step 3: Commit**

```bash
git add src/reader.rs
git commit -m "feat: add sheet enumeration helpers"
```

---

## Task 5: CSV Writer Module

**Files:**
- Create: `src/writer.rs`
- Modify: `src/main.rs`

**Step 1: Create src/writer.rs**

```rust
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
            let value = sheet.get_formatted_value((col, row));
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
```

**Step 2: Add module to main.rs**

Add after `mod reader;`:

```rust
mod writer;
```

**Step 3: Verify it compiles**

Run: `cargo build`
Expected: Compiles successfully

**Step 4: Commit**

```bash
git add src/writer.rs src/main.rs
git commit -m "feat: add CSV writer module"
```

---

## Task 6: Main Integration - Single Sheet to Stdout

**Files:**
- Modify: `src/main.rs`

**Step 1: Update run function for basic conversion**

Replace the `run` function in `src/main.rs`:

```rust
use std::io::{self, Write};

fn run(args: Args) -> error::Result<()> {
    if args.verbose {
        eprintln!("reading: {:?}", args.input);
    }

    let workbook = reader::open_workbook(&args.input)?;
    let sheet_names = reader::get_sheet_names(&workbook);

    if args.verbose {
        eprintln!("format: {}", args.input.extension().unwrap_or_default().to_string_lossy());
        eprintln!("sheets: {}", sheet_names.join(", "));
    }

    let config = writer::CsvConfig {
        format: args.format,
        empty_value: args.empty.clone(),
    };

    // Determine which sheets to process
    let sheets_to_process: Vec<usize> = match &args.sheet {
        Some(selector) => {
            let sel = reader::parse_sheet_selector(selector);
            vec![resolve_sheet_index(&workbook, &sel, &sheet_names)?]
        }
        None => (0..sheet_names.len()).collect(),
    };

    // Check output constraints
    if sheets_to_process.len() > 1 && args.output.is_none() {
        return Err(error::Error::MultipleSheetNoOutput);
    }

    // Process sheets
    for &sheet_idx in &sheets_to_process {
        let sheet = workbook
            .get_sheet(&(sheet_idx as u32))
            .ok_or_else(|| error::Error::SheetIndexOutOfRange {
                index: sheet_idx,
                count: sheet_names.len(),
            })?;

        if args.verbose {
            let (cols, rows) = sheet.get_highest_column_and_row();
            eprintln!("processing: {} ({} rows, {} columns)", sheet.get_name(), rows, cols);
        }

        match &args.output {
            None => {
                // Single sheet to stdout
                let stdout = io::stdout();
                let handle = stdout.lock();
                writer::write_sheet(sheet, handle, &config)?;
            }
            Some(output_path) => {
                let file_path = if output_path.is_dir() || sheets_to_process.len() > 1 {
                    // Output to directory
                    let dir = if output_path.exists() && output_path.is_dir() {
                        output_path.clone()
                    } else {
                        std::fs::create_dir_all(output_path)?;
                        output_path.clone()
                    };
                    dir.join(format!("{}.csv", sheet.get_name()))
                } else {
                    output_path.clone()
                };

                if args.verbose {
                    eprintln!("output: {:?}", file_path);
                }

                let file = std::fs::File::create(&file_path)?;
                writer::write_sheet(sheet, file, &config)?;
            }
        }
    }

    Ok(())
}

fn resolve_sheet_index(
    workbook: &umya_spreadsheet::Spreadsheet,
    selector: &reader::SheetSelector,
    sheet_names: &[String],
) -> error::Result<usize> {
    match selector {
        reader::SheetSelector::Index(idx) => {
            if *idx < sheet_names.len() {
                Ok(*idx)
            } else {
                Err(error::Error::SheetIndexOutOfRange {
                    index: *idx,
                    count: sheet_names.len(),
                })
            }
        }
        reader::SheetSelector::Name(name) => {
            sheet_names
                .iter()
                .position(|n| n == name)
                .ok_or_else(|| error::Error::SheetNotFound {
                    name: name.clone(),
                    available: sheet_names.join(", "),
                })
        }
    }
}
```

**Step 2: Add missing import**

Add at top of main.rs after other imports:

```rust
use std::io::{self, Write};
```

**Step 3: Verify it compiles**

Run: `cargo build`
Expected: Compiles successfully

**Step 4: Commit**

```bash
git add src/main.rs
git commit -m "feat: integrate reader and writer for full conversion"
```

---

## Task 7: Manual Testing with Real Files

**Files:** None (testing only)

**Step 1: Create a test XLSX file**

You'll need a test Excel file. Create one manually or download a sample.

**Step 2: Test basic conversion**

Run: `cargo run -- test.xlsx`
Expected: CSV output to stdout

**Step 3: Test with output file**

Run: `cargo run -- test.xlsx -o output.csv`
Expected: Creates output.csv

**Step 4: Test verbose mode**

Run: `cargo run -- test.xlsx -v`
Expected: Progress info to stderr, CSV to stdout

**Step 5: Test format options**

Run: `cargo run -- test.xlsx -f tsv`
Expected: Tab-separated output

**Step 6: Test sheet selection**

Run: `cargo run -- test.xlsx -s 0`
Expected: First sheet only

---

## Task 8: Error Handling Polish

**Files:**
- Modify: `src/main.rs`

**Step 1: Improve file not found message**

The current implementation should already handle this via `Error::FileNotFound`.

**Step 2: Test error cases**

Run: `cargo run -- nonexistent.xlsx`
Expected: `error: file not found: nonexistent.xlsx` and exit code 1

Run: `cargo run -- Cargo.toml`
Expected: `error: unsupported file format: toml` and exit code 3

**Step 3: Verify exit codes**

Run: `cargo run -- nonexistent.xlsx; echo "Exit code: $?"`
Expected: Exit code: 1

**Step 4: Commit if any changes**

```bash
git add -A
git commit -m "fix: polish error handling" --allow-empty
```

---

## Task 9: Build Release and Final Test

**Files:** None

**Step 1: Build release binary**

Run: `cargo build --release`
Expected: Binary at `target/release/excel2csv`

**Step 2: Test release binary**

Run: `./target/release/excel2csv --help`
Expected: Help output displays correctly

**Step 3: Test with actual Excel file**

Run: `./target/release/excel2csv test.xlsx`
Expected: Formatted CSV output (currency shows as `$1,234.56`, dates formatted)

**Step 4: Final commit**

```bash
git add -A
git commit -m "chore: release build verified"
```

---

## Summary

After completing all tasks you'll have:

1. A working `excel2csv` CLI tool
2. Support for both .xls and .xlsx files
3. Formatted cell output (currency, dates, etc.)
4. Multi-sheet handling with `-s` selector
5. Output format presets (csv, tsv, european)
6. Configurable empty cell values
7. Verbose mode for debugging
8. Proper error messages and exit codes
