// The Assumptions sheet: DemoCo's forecast drivers per year and the 2024A
// revenue base, as plain unformatted-font inputs for Autocolor to find.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{
    assumptions::*, last_year_col, year_col, FIRST_YEAR_COL, HEADER_ROW, LABEL_COL, TITLE_ROW,
    UNIT_COL, YEARS,
};
use crate::pen::Pen;
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = crate::layout::ASSUMPTIONS_SHEET;
const NOTE_ROW: u32 = 1;
const FORECAST_YEARS: usize = YEARS.len() - 1;
const LABEL_WIDTH: f64 = 30.0;
const UNIT_WIDTH: f64 = 7.0;
const YEAR_WIDTH: f64 = 11.0;
const REVENUE_2024A: f64 = 12_400.0;

struct Driver {
    row: u32,
    label: &'static str,
    unit: &'static str,
    values: [f64; FORECAST_YEARS],
    percent: bool,
}

const DRIVERS: [Driver; 6] = [
    Driver { row: GROWTH, label: "Revenue growth", unit: "%", values: [0.08, 0.07, 0.06, 0.05, 0.05], percent: true },
    Driver { row: GROSS_MARGIN, label: "Gross margin", unit: "%", values: [0.38, 0.385, 0.39, 0.39, 0.395], percent: true },
    Driver { row: OPEX_RATIO, label: "Operating expenses, % of revenue", unit: "%", values: [0.22, 0.215, 0.21, 0.21, 0.205], percent: true },
    Driver { row: DA_RATIO, label: "D&A, % of revenue", unit: "%", values: [0.035, 0.035, 0.035, 0.035, 0.035], percent: true },
    Driver { row: INTEREST, label: "Interest expense", unit: "EUR k", values: [180.0, 170.0, 160.0, 150.0, 140.0], percent: false },
    Driver { row: TAX_RATE, label: "Tax rate", unit: "%", values: [0.2, 0.2, 0.2, 0.2, 0.2], percent: true },
];

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "DemoCo SIA: model assumptions (EUR thousands)", &styles.title)?;
    pen.text(NOTE_ROW, LABEL_COL, "Inputs are plain numbers on purpose: run Autocolor and they turn blue.", &styles.note)?;
    pen.text(HEADER_ROW, LABEL_COL, "Driver", &styles.header_left)?;
    pen.text(HEADER_ROW, UNIT_COL, "Unit", &styles.header_left)?;
    for (i, year) in YEARS[1..].iter().enumerate() {
        pen.text(HEADER_ROW, year_col(i), year, &styles.header)?;
    }
    for driver in &DRIVERS {
        let format = if driver.percent { &styles.pct } else { &styles.eur };
        pen.text(driver.row, LABEL_COL, driver.label, &styles.label)?;
        pen.text(driver.row, UNIT_COL, driver.unit, &styles.label)?;
        for (i, value) in driver.values.iter().enumerate() {
            pen.number(driver.row, year_col(i), *value, format)?;
        }
    }
    pen.text(REVENUE_2024, LABEL_COL, "Revenue 2024A (actual)", &styles.label)?;
    pen.text(REVENUE_2024, UNIT_COL, "EUR k", &styles.label)?;
    pen.number(REVENUE_2024, FIRST_YEAR_COL, REVENUE_2024A, &styles.eur)?;
    let sheet = pen.sheet();
    sheet.set_column_width(LABEL_COL, LABEL_WIDTH)?;
    sheet.set_column_width(UNIT_COL, UNIT_WIDTH)?;
    sheet.set_column_range_width(FIRST_YEAR_COL, last_year_col(), YEAR_WIDTH)?;
    Ok(pen.tally)
}
