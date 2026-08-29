// The tornado input: a Driver | Low | High block with the heading and the base
// case in the row above it, exactly the shape the Tornado button reads, and an
// empty block to its right for the helper table the tool writes. The swing
// sizes live in their own input table below, so the block holds no hardcodes.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{
    assumptions as a, assumption_col, assumptions_ref_abs, cell, cell_abs, pnl::*, pnl_ref_abs,
    year_col, LABEL_COL, TITLE_ROW,
};
use crate::pen::Pen;
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = "Sensitivity";
const NOTE_ROW: u32 = 1;
const BASE_ROW: u32 = 2;
const HEADER_ROW: u32 = 3;
const FIRST_DRIVER_ROW: u32 = 4;
const DRIVER_COUNT: u32 = 5;
const DRIVER_COL: u16 = 0;
const LOW_COL: u16 = 1;
const HIGH_COL: u16 = 2;
const SWINGS_TITLE_ROW: u32 = 10;
const FIRST_SWING_ROW: u32 = 11;
const SWING_VALUE_COL: u16 = 1;
const DRIVER_WIDTH: f64 = 40.0;
const VALUE_WIDTH: f64 = 11.0;
/// The P&L year the sensitivity is run on (2026E).
const YEAR: usize = 2;
/// Two forecast years of growth move before 2026E.
const GROWTH_YEARS: &str = "2";

const SWINGS: [(&str, f64); 6] = [
    ("Price move", 0.03),
    ("Growth move, pp per year", 0.02),
    ("Gross margin move, pp", 0.01),
    ("Opex ratio move, pp", 0.015),
    ("SEK rate move", 0.05),
    ("SEK share of sales", 0.25),
];

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "EBITDA 2026E sensitivity (EUR thousands)", &styles.title)?;
    pen.text(NOTE_ROW, LABEL_COL, &format!("Select {} and press Tornado: the heading and the base case come from row {}.", block_address(), BASE_ROW + 1), &styles.note)?;
    pen.text(BASE_ROW, DRIVER_COL, "EBITDA 2026E sensitivity, EUR k", &styles.label_bold)?;
    pen.formula(BASE_ROW, LOW_COL, &format!("={}", pnl_ref_abs(EBITDA, year_col(YEAR))), &styles.eur_bold)?;
    pen.text(HEADER_ROW, DRIVER_COL, "Driver", &styles.header_left)?;
    pen.text(HEADER_ROW, LOW_COL, "Low", &styles.header)?;
    pen.text(HEADER_ROW, HIGH_COL, "High", &styles.header)?;
    for (i, (label, delta)) in drivers().iter().enumerate() {
        let row = FIRST_DRIVER_ROW + i as u32;
        pen.text(row, DRIVER_COL, label, &styles.label)?;
        pen.formula(row, LOW_COL, &format!("={}-{delta}", base()), &styles.eur)?;
        pen.formula(row, HIGH_COL, &format!("={}+{delta}", base()), &styles.eur)?;
    }
    pen.text(SWINGS_TITLE_ROW, DRIVER_COL, "Swing sizes (inputs)", &styles.label_bold)?;
    for (i, (label, value)) in SWINGS.iter().enumerate() {
        let row = FIRST_SWING_ROW + i as u32;
        pen.text(row, DRIVER_COL, label, &styles.label)?;
        pen.number(row, SWING_VALUE_COL, *value, &styles.pct)?;
    }
    let sheet = pen.sheet();
    sheet.set_column_width(DRIVER_COL, DRIVER_WIDTH)?;
    sheet.set_column_range_width(LOW_COL, HIGH_COL, VALUE_WIDTH)?;
    Ok(pen.tally)
}

/// A1 address of the Driver | Low | High block, header row included.
pub fn block_address() -> String {
    format!("{}:{}", cell(HEADER_ROW, DRIVER_COL), cell(FIRST_DRIVER_ROW + DRIVER_COUNT - 1, HIGH_COL))
}

fn base() -> String {
    cell_abs(BASE_ROW, LOW_COL)
}

fn swing(i: u32) -> String {
    cell_abs(FIRST_SWING_ROW + i, SWING_VALUE_COL)
}

/// Each driver's EBITDA delta, symmetric around the base.
fn drivers() -> [(&'static str, String); DRIVER_COUNT as usize] {
    let revenue = pnl_ref_abs(REVENUE, year_col(YEAR));
    let margin = pnl_ref_abs(GROSS_MARGIN, year_col(YEAR));
    let opex_ratio = assumptions_ref_abs(a::OPEX_RATIO, assumption_col(YEAR));
    [
        ("Price -3% / +3%", format!("{}*{revenue}", swing(0))),
        (
            "Revenue growth -2pp / +2pp, two years",
            format!("{GROWTH_YEARS}*{}*{revenue}*({margin}-{opex_ratio})", swing(1)),
        ),
        ("Gross margin -1pp / +1pp", format!("{}*{revenue}", swing(2))),
        ("Opex ratio +1.5pp / -1.5pp", format!("{}*{revenue}", swing(3))),
        ("SEK rate -5% / +5% on SEK sales", format!("{}*{}*{revenue}*{margin}", swing(4), swing(5))),
    ]
}

/// A1 address of the base-case cell above the block.
pub fn base_address() -> String {
    cell(BASE_ROW, LOW_COL)
}
