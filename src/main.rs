mod error;
mod reader;
mod writer;

use clap::{Parser, ValueEnum};
use std::io;
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
            .get_sheet(&sheet_idx)
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
    _workbook: &umya_spreadsheet::Spreadsheet,
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
