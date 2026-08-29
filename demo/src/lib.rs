// pls,fix demo workbook: a small DemoCo model whose every sheet gives
// one of the add-in's tools something to act on. `build` returns the workbook
// and a per-sheet tally of what was written, so main.rs can log it and the
// test can reconcile the saved file against it, cell for cell.

pub mod layout;
pub mod pen;
pub mod sheets;
pub mod style;
pub mod tally;

use rust_xlsxwriter::{Workbook, XlsxError};

use crate::style::Styles;
use crate::tally::Tally;

/// Builds the whole workbook; sheets in `sheets::SHEETS` order.
pub fn build() -> Result<(Workbook, Vec<(&'static str, Tally)>), XlsxError> {
    let styles = Styles::new();
    let mut workbook = Workbook::new();
    let mut tallies = Vec::with_capacity(sheets::SHEETS.len());
    for (name, builder) in sheets::SHEETS {
        let sheet = workbook.add_worksheet();
        sheet.set_name(name)?;
        let tally = builder(sheet, &styles)?;
        tallies.push((name, tally));
    }
    sheets::define_names(&mut workbook)?;
    Ok((workbook, tallies))
}
