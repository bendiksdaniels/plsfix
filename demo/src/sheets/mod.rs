// The sheet builders in workbook order, plus the defined names that span
// sheets. Each builder owns one sheet and returns what it wrote.

pub mod assumptions;
pub mod bridge;
pub mod data;
pub mod pnl;
pub mod rounding;
pub mod scratch;
pub mod sensitivity;
pub mod start;
pub mod variance;

use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};

use crate::layout::{assumptions as a, assumptions_ref_abs, FIRST_YEAR_COL};
use crate::style::Styles;
use crate::tally::Tally;

pub type Builder = fn(&mut Worksheet, &Styles) -> Result<Tally, XlsxError>;

pub const SHEETS: [(&str, Builder); 9] = [
    (start::NAME, start::build as Builder),
    (assumptions::NAME, assumptions::build as Builder),
    (pnl::NAME, pnl::build as Builder),
    (bridge::NAME, bridge::build as Builder),
    (sensitivity::NAME, sensitivity::build as Builder),
    (rounding::NAME, rounding::build as Builder),
    (variance::NAME, variance::build as Builder),
    (data::NAME, data::build as Builder),
    (scratch::NAME, scratch::build as Builder),
];

/// The name of a defined name that points nowhere: what Scan broken names reports.
pub const BROKEN_NAME: &str = "Old_budget";
const BROKEN_FORMULA: &str = "=#REF!";

/// Two names the model uses, one nothing uses, one left broken on purpose.
pub fn define_names(workbook: &mut Workbook) -> Result<(), XlsxError> {
    let revenue = format!("={}", assumptions_ref_abs(a::REVENUE_2024, FIRST_YEAR_COL));
    let tax = format!("={}", assumptions_ref_abs(a::TAX_RATE, FIRST_YEAR_COL));
    workbook.define_name("Revenue_2024A", &revenue)?;
    workbook.define_name("TaxRate_2025E", &tax)?;
    workbook.define_name(scratch::AREA_NAME, &scratch::area_formula())?;
    workbook.define_name(BROKEN_NAME, BROKEN_FORMULA)?;
    Ok(())
}
