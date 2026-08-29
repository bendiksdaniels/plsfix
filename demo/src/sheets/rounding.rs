// Consistent rounding: a segment split whose Excel-rounded shares add up to
// 101, the case Consistent rounding and =PLSFIX.ROUND / =PLSFIX.ROUNDSUM exist for.
// The column right of the share points stays empty: the tool writes there.
// The custom-function formulas are shown as text to type (without the leading
// =, which the sharing scan would otherwise count as a formula), so the file
// opens clean on a machine without the add-in.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{cell, cell_abs, column, LABEL_COL, TITLE_ROW};
use crate::pen::Pen;
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = "Rounding";
const NOTE_ROW: u32 = 1;
const HEADER_ROW: u32 = 2;
const FIRST_SEGMENT_ROW: u32 = 3;
const SEGMENT_COUNT: u32 = 5;
const TOTAL_ROW: u32 = 8;
const HINT_ROW: u32 = 10;
const SEGMENT_COL: u16 = 0;
const REVENUE_COL: u16 = 1;
const POINTS_COL: u16 = 2;
const LANDING_COL: u16 = 3;
const EXCEL_COL: u16 = 4;
const TYPE_IT_COL: u16 = 5;
const SEGMENT_WIDTH: f64 = 16.0;
const VALUE_WIDTH: f64 = 15.0;
const LANDING_WIDTH: f64 = 26.0;
const TYPE_IT_WIDTH: f64 = 32.0;
const DECIMALS: &str = "0";
const POINTS_PER_UNIT: &str = "100";

const SEGMENTS: [(&str, f64); SEGMENT_COUNT as usize] = [
    ("Retail", 3_162.0),
    ("Wholesale", 3_162.0),
    ("Export", 3_038.0),
    ("Food service", 1_519.0),
    ("Online", 1_519.0),
];

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "Consistent rounding (the think-cell TCROUND idea)", &styles.title)?;
    pen.text(NOTE_ROW, LABEL_COL, &format!("Select {} and press Consistent rounding, or type the formulas shown in column {}.", points_address(), column(TYPE_IT_COL)), &styles.note)?;
    write_header(&mut pen, styles)?;
    for (i, (segment, revenue)) in SEGMENTS.iter().enumerate() {
        let row = FIRST_SEGMENT_ROW + i as u32;
        pen.text(row, SEGMENT_COL, segment, &styles.label)?;
        pen.number(row, REVENUE_COL, *revenue, &styles.eur)?;
        pen.formula(row, POINTS_COL, &format!("={}/{}*{POINTS_PER_UNIT}", cell(row, REVENUE_COL), cell_abs(TOTAL_ROW, REVENUE_COL)), &styles.points)?;
        pen.formula(row, EXCEL_COL, &format!("=ROUND({},{DECIMALS})", cell(row, POINTS_COL)), &styles.whole)?;
        pen.text(row, TYPE_IT_COL, &format!("PLSFIX.ROUND({}, {}, {DECIMALS})", points_address_abs(), i + 1), &styles.plain)?;
    }
    write_total(&mut pen, styles)?;
    pen.text(HINT_ROW, LABEL_COL, &format!("{} shows 101 although the shares sum to 100: consistent rounding gives parts that add up.", cell(TOTAL_ROW, EXCEL_COL)), &styles.note)?;
    let sheet = pen.sheet();
    sheet.set_column_width(SEGMENT_COL, SEGMENT_WIDTH)?;
    sheet.set_column_range_width(REVENUE_COL, POINTS_COL, VALUE_WIDTH)?;
    sheet.set_column_width(LANDING_COL, LANDING_WIDTH)?;
    sheet.set_column_width(EXCEL_COL, VALUE_WIDTH)?;
    sheet.set_column_width(TYPE_IT_COL, TYPE_IT_WIDTH)?;
    Ok(pen.tally)
}

/// A1 address of the share-points column, the Consistent rounding input.
pub fn points_address() -> String {
    format!("{}:{}", cell(FIRST_SEGMENT_ROW, POINTS_COL), cell(TOTAL_ROW - 1, POINTS_COL))
}

fn points_address_abs() -> String {
    format!("{}:{}", cell_abs(FIRST_SEGMENT_ROW, POINTS_COL), cell_abs(TOTAL_ROW - 1, POINTS_COL))
}

fn write_header(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    pen.text(HEADER_ROW, SEGMENT_COL, "Segment", &styles.header_left)?;
    pen.text(HEADER_ROW, REVENUE_COL, "Revenue, EUR k", &styles.header)?;
    pen.text(HEADER_ROW, POINTS_COL, "Share, points", &styles.header)?;
    pen.text(HEADER_ROW, LANDING_COL, "Consistent rounding lands here", &styles.header_left)?;
    pen.text(HEADER_ROW, EXCEL_COL, "Excel ROUND", &styles.header)?;
    pen.text(HEADER_ROW, TYPE_IT_COL, "Type these with = in front", &styles.header_left)?;
    Ok(())
}

fn write_total(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    let last = TOTAL_ROW - 1;
    pen.text(TOTAL_ROW, SEGMENT_COL, "Total", &styles.label_bold)?;
    for (col, format) in [(REVENUE_COL, &styles.eur_bold), (POINTS_COL, &styles.points), (EXCEL_COL, &styles.whole)] {
        pen.formula(TOTAL_ROW, col, &format!("=SUM({}:{})", cell(FIRST_SEGMENT_ROW, col), cell(last, col)), format)?;
    }
    pen.text(TOTAL_ROW, TYPE_IT_COL, &format!("PLSFIX.ROUNDSUM({}, {DECIMALS})", points_address_abs()), &styles.plain)?;
    Ok(())
}
