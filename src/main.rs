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
