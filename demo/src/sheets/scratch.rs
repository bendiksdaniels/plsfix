// A hidden sheet of stale figures with a defined name nothing uses: what
// Prepare for sharing and the names scan are there to report.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{cell_abs, LABEL_COL};
use crate::pen::Pen;
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = "Scratch";
pub const AREA_NAME: &str = "Scratch_area";
const FIRST_ROW: u32 = 0;
const VALUE_COL: u16 = 1;
const LABEL_WIDTH: f64 = 60.0;

const ROWS: [(&str, f64); 6] = [
    ("Scratch: stale figures, hidden on purpose so Prepare for sharing has something to find", 0.0),
    ("Old revenue 2024 budget", 12_100.0),
    ("Old EBITDA 2024 budget", 1_950.0),
    ("Old headcount", 184.0),
    ("Old capex", 640.0),
    ("Old net debt", 3_300.0),
];

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut pen = Pen::new(sheet);
    for (i, (label, value)) in ROWS.iter().enumerate() {
        let row = FIRST_ROW + i as u32;
        pen.text(row, LABEL_COL, label, &styles.label)?;
        pen.number(row, VALUE_COL, *value, &styles.eur)?;
    }
    let sheet = pen.sheet();
    sheet.set_column_width(LABEL_COL, LABEL_WIDTH)?;
    sheet.set_hidden(true);
    Ok(pen.tally)
}

/// The workbook-level formula behind `Scratch_area`.
pub fn area_formula() -> String {
    format!(
        "={NAME}!{}:{}",
        cell_abs(FIRST_ROW, LABEL_COL),
        cell_abs(FIRST_ROW + ROWS.len() as u32 - 1, VALUE_COL)
    )
}
