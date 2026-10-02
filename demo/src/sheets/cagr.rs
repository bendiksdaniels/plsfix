// The CAGR sheet: six series that put the ribbon's Insert CAGR, the
// =PLSFIX.CAGR custom function and a plain hand formula side by side. Four
// series should agree on one rate; a zero or negative starting value should
// make all three refuse instead of guessing.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{a1_row, cell, range, GUIDE_ROWS, LABEL_COL, TITLE_ROW};
use crate::pen::Pen;
use crate::sheets::guide::{self, Guide};
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = "CAGR";
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const HEADER_ROW: u32 = GUIDE_ROWS + 2;
const FIRST_ROW: u32 = GUIDE_ROWS + 3;
const SERIES_COUNT: u32 = 6;
/// The first four series (growing, falling, flat, volatile) are where the
/// ribbon, the function and the hand formula all agree on one rate; the last
/// two (zero and negative start) should make every method refuse instead.
const AGREE_COUNT: u32 = 4;
const EXPECTED_ROW: u32 = FIRST_ROW + SERIES_COUNT;
const GUIDE: Guide = Guide {
    title: "three ways to check a five-year CAGR",
    tasks: &[
        "Select a series' six years and press Insert CAGR on the pls,fix ribbon: the result lands right of the row.",
        "Column I uses =PLSFIX.CAGR(first, last, 5), column J the hand formula; the three must agree.",
    ],
    commands: &["Insert CAGR (Ctrl+Shift+Q)"],
};
const _: () = assert!(GUIDE.tasks.len() <= guide::MAX_TASKS, "a fifth task would land on the Commands row");

const SERIES_COL: u16 = 0;
const FIRST_YEAR_COL: u16 = 1;
const LAST_YEAR_COL: u16 = 6;
const RIBBON_COL: u16 = 7;
const FUNCTION_COL: u16 = 8;
const HAND_COL: u16 = 9;
const PERIODS: u32 = 5;

const SERIES_WIDTH: f64 = 30.0;
const YEAR_WIDTH: f64 = 10.0;
const RESULT_WIDTH: f64 = 14.0;

const YEARS: [&str; 6] = ["2020", "2021", "2022", "2023", "2024", "2025"];

/// (label, six years of value): the four series that should agree, then the
/// two that should make every method refuse.
const SERIES: [(&str, [f64; 6]); SERIES_COUNT as usize] = [
    ("Revenue, +10% a year", [100.0, 110.0, 121.0, 133.1, 146.41, 161.051]),
    ("Costs, falling", [80.0, 76.0, 71.0, 66.0, 63.0, 60.0]),
    ("Users, flat", [1000.0, 1050.0, 990.0, 1020.0, 1000.0, 1000.0]),
    ("Volatile (first and last count)", [50.0, 120.0, 30.0, 90.0, 40.0, 110.0]),
    ("Zero start (should refuse)", [0.0, 10.0, 20.0, 30.0, 40.0, 50.0]),
    ("Negative start (should refuse)", [-20.0, -10.0, 0.0, 10.0, 20.0, 30.0]),
];

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "CAGR check: ribbon, function and hand formula", &styles.title)?;
    pen.text(
        NOTE_ROW,
        LABEL_COL,
        &format!(
            "Select {} and press Insert CAGR; column I (=PLSFIX.CAGR) and column J (hand formula) should land on the same rate, or all three should refuse together.",
            first_row_range()
        ),
        &styles.note,
    )?;
    write_header(&mut pen, styles)?;
    for (i, (label, values)) in SERIES.iter().enumerate() {
        let row = FIRST_ROW + i as u32;
        write_series_row(&mut pen, styles, row, label, values)?;
    }
    pen.text(EXPECTED_ROW, LABEL_COL, "Expected: 10.0%, -5.6%, 0.0%, 17.1%, refused, refused.", &styles.note)?;
    let sheet = pen.sheet();
    sheet.set_column_width(SERIES_COL, SERIES_WIDTH)?;
    sheet.set_column_range_width(FIRST_YEAR_COL, LAST_YEAR_COL, YEAR_WIDTH)?;
    sheet.set_column_range_width(RIBBON_COL, HAND_COL, RESULT_WIDTH)?;
    // Column H carries the percentage format even though no cell is written
    // there yet, so the ribbon's own write (outside this generator) lands
    // already formatted instead of General.
    sheet.set_column_format(RIBBON_COL, &styles.pct)?;
    tally += pen.tally;
    Ok(tally)
}

fn write_header(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    pen.text(HEADER_ROW, SERIES_COL, "Series", &styles.header_left)?;
    for (i, year) in YEARS.iter().enumerate() {
        pen.text(HEADER_ROW, FIRST_YEAR_COL + i as u16, year, &styles.header)?;
    }
    pen.text(HEADER_ROW, RIBBON_COL, "Ribbon CAGR", &styles.header)?;
    pen.text(HEADER_ROW, FUNCTION_COL, "=PLSFIX.CAGR", &styles.header)?;
    pen.text(HEADER_ROW, HAND_COL, "Hand formula", &styles.header)?;
    Ok(())
}

/// One series: its label, the six years, then the function and hand formulas.
/// Column RIBBON_COL (H) is left empty on purpose: the ribbon button writes there.
fn write_series_row(pen: &mut Pen, styles: &Styles, row: u32, label: &str, values: &[f64; 6]) -> Result<(), XlsxError> {
    pen.text(row, SERIES_COL, label, &styles.label)?;
    for (y, value) in values.iter().enumerate() {
        pen.number(row, FIRST_YEAR_COL + y as u16, *value, &styles.number_1dp)?;
    }
    let first = cell(row, FIRST_YEAR_COL);
    let last = cell(row, LAST_YEAR_COL);
    // Excel stores a custom function's formula under its own prefix,
    // _xldudf_<namespace>_<name>(...), once the workbook is saved (proven on
    // Excel for Mac 16.107). A plain "PLSFIX.CAGR(...)" written here instead
    // is never resolved and the cell shows #NAME? the moment the file opens,
    // so this writes the exact form Excel itself would have written.
    pen.formula(row, FUNCTION_COL, &format!("=_xldudf_PLSFIX_CAGR({first},{last},{PERIODS})"), &styles.pct)?;
    pen.formula(row, HAND_COL, &format!("=IFERROR(({last}/{first})^(1/{PERIODS})-1,\"n/a\")"), &styles.pct)?;
    Ok(())
}

/// A1 range of the first series' six years: the ribbon CAGR's example input.
pub fn first_row_range() -> String {
    range(FIRST_ROW, FIRST_YEAR_COL, FIRST_ROW, LAST_YEAR_COL)
}

/// "rows N to M": the series where the ribbon, the function and the hand
/// formula all land on the same rate.
pub fn agree_rows_words() -> String {
    format!("rows {} to {}", a1_row(FIRST_ROW), a1_row(FIRST_ROW + AGREE_COUNT - 1))
}

/// "rows N and M": the series (zero and negative start) every method refuses.
pub fn refused_rows_words() -> String {
    format!("rows {} and {}", a1_row(FIRST_ROW + AGREE_COUNT), a1_row(FIRST_ROW + SERIES_COUNT - 1))
}
