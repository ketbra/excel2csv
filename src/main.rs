mod error;
mod evaluator;
mod reader;
mod writer;

use clap::{Parser, ValueEnum};
use std::collections::HashMap;
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

    /// Recalculate all formulas before export
    #[arg(short, long)]
    pub recalculate: bool,

    /// Override cell value before recalculation (format: A1=value or Sheet1!A1=value)
    /// Can be specified multiple times. Implies --recalculate.
    #[arg(long = "set", value_name = "CELL=VALUE")]
    pub set_cells: Vec<String>,

    /// Override TODAY()/NOW() date for reproducible output (format: YYYY-MM-DD)
    #[arg(long)]
    pub today: Option<String>,
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

    // Determine if we need to recalculate (explicit flag or --set provided)
    let needs_recalc = args.recalculate || !args.set_cells.is_empty();

    // Parse --set arguments into a HashMap
    let overrides: HashMap<String, String> = args
        .set_cells
        .iter()
        .filter_map(|s| {
            let parts: Vec<&str> = s.splitn(2, '=').collect();
            if parts.len() == 2 {
                Some((parts[0].to_string(), parts[1].to_string()))
            } else {
                eprintln!("warning: ignoring invalid --set format: {}", s);
                None
            }
        })
        .collect();

    // Parse --today
    let today = args.today.as_ref().and_then(|s| {
        chrono::NaiveDate::parse_from_str(s, "%Y-%m-%d")
            .map_err(|e| eprintln!("warning: invalid --today format: {} ({})", s, e))
            .ok()
    });

    // Load evaluated workbook if recalculation is needed
    let evaluated = if needs_recalc {
        if args.verbose {
            eprintln!("recalculating formulas...");
            if !overrides.is_empty() {
                eprintln!("  overrides: {:?}", overrides);
            }
            if let Some(ref t) = today {
                eprintln!("  today: {}", t);
            }
        }

        let eval_config = evaluator::EvalConfig { overrides, today };
        Some(evaluator::EvaluatedWorkbook::load(&args.input, &eval_config)?)
    } else {
        None
    };

    // Get sheet names from evaluated workbook or regular workbook
    let (workbook, sheet_names) = if let Some(ref eval_wb) = evaluated {
        (None, eval_wb.sheet_names().to_vec())
    } else {
        let wb = reader::open_workbook(&args.input)?;
        let names = reader::get_sheet_names(&wb);
        (Some(wb), names)
    };

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
            vec![resolve_sheet_index_from_names(&sel, &sheet_names)?]
        }
        None => (0..sheet_names.len()).collect(),
    };

    // Check output constraints
    if sheets_to_process.len() > 1 && args.output.is_none() {
        return Err(error::Error::MultipleSheetNoOutput);
    }

    // Process sheets
    for &sheet_idx in &sheets_to_process {
        let sheet_name = &sheet_names[sheet_idx];

        if args.verbose {
            eprintln!("processing: {}", sheet_name);
        }

        match &args.output {
            None => {
                // Single sheet to stdout
                let stdout = io::stdout();
                let handle = stdout.lock();
                if let Some(ref eval_wb) = evaluated {
                    writer::write_evaluated_sheet(eval_wb, sheet_name, handle, &config)?;
                } else {
                    let wb = workbook.as_ref().unwrap();
                    let sheet = wb.get_sheet(&sheet_idx).ok_or(error::Error::SheetIndexOutOfRange {
                        index: sheet_idx,
                        count: sheet_names.len(),
                    })?;
                    writer::write_sheet(sheet, handle, &config)?;
                }
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
                    dir.join(format!("{}.csv", sheet_name))
                } else {
                    output_path.clone()
                };

                if args.verbose {
                    eprintln!("output: {:?}", file_path);
                }

                let file = std::fs::File::create(&file_path)?;
                if let Some(ref eval_wb) = evaluated {
                    writer::write_evaluated_sheet(eval_wb, sheet_name, file, &config)?;
                } else {
                    let wb = workbook.as_ref().unwrap();
                    let sheet = wb.get_sheet(&sheet_idx).ok_or(error::Error::SheetIndexOutOfRange {
                        index: sheet_idx,
                        count: sheet_names.len(),
                    })?;
                    writer::write_sheet(sheet, file, &config)?;
                }
            }
        }
    }

    Ok(())
}

fn resolve_sheet_index_from_names(
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
