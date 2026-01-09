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
