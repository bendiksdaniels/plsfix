// The Variance sheet: a small bank-deposit reconciliation for "Find a
// combination". Fourteen invoice lines, mixed sign; a Target that exactly
// one subset reaches. tests/workbook.rs proves that by brute force, so the
// puzzle can never quietly grow a second solution or lose its only one.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{
    cell, range, variance::*, GUIDE_ROWS, HEADER_ROW, LABEL_COL, TITLE_ROW, VARIANCE_SHEET,
};
use crate::pen::Pen;
use crate::sheets::guide::{self, Guide};
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = VARIANCE_SHEET;
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const GUIDE: Guide = Guide {
    title: "a bank deposit variance to reconcile",
    tasks: &["Find a combination that reaches the Target.", "Comps stats runs under the table."],
    commands: &["Open pls,fix (Ctrl+Shift+M)"],
};
const LABEL_WIDTH: f64 = 38.0;
const AMOUNT_WIDTH: f64 = 14.0;

/// (label, amount): ten invoices and four credit notes, mixed sign.
pub const AMOUNTS: [(&str, f64); COUNT as usize] = [
    ("Invoice 1042, Rimi Retail", 2_480.0),
    ("Invoice 1043, Maxima Wholesale", 3_150.0),
    ("Credit note 118, Rimi Retail", -420.0),
    ("Invoice 1044, Narvesen kiosk", 610.0),
    ("Invoice 1045, Circle K cafe", 875.0),
    ("Credit note 119, Maxima Wholesale", -390.0),
    ("Invoice 1046, Hotel Bergs", 1_920.0),
    ("Invoice 1047, Riga airport catering", 3_340.0),
    ("Credit note 120, Hotel Bergs", -260.0),
    ("Invoice 1048, Elvi Latvia", 1_180.0),
    ("Invoice 1049, Depo wholesale", 2_760.0),
    ("Credit note 121, Elvi Latvia", -150.0),
    ("Invoice 1050, Rimi Retail", 990.0),
    ("Invoice 1051, Narvesen kiosk", 705.0),
];

/// The stated target: exactly one subset of AMOUNTS reaches it (proved in
/// tests/workbook.rs), never written into the sheet itself.
pub const TARGET: f64 = 2_230.0;

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "DemoCo SIA: bank deposit variance", &styles.title)?;
    pen.text(
        NOTE_ROW,
        LABEL_COL,
        &format!(
            "A bank deposit landed with no reference. Select {} and type the number from {} into the pane's Find a combination, then press Find cells.",
            amounts_address(),
            target_address()
        ),
        &styles.note,
    )?;
    pen.text(HEADER_ROW, LABEL_COL, "Line item", &styles.header_left)?;
    pen.text(HEADER_ROW, AMOUNT_COL, "Amount, EUR", &styles.header)?;
    for (i, (label, amount)) in AMOUNTS.iter().enumerate() {
        let row = FIRST_ROW + i as u32;
        pen.text(row, LABEL_COL, label, &styles.label)?;
        pen.number(row, AMOUNT_COL, *amount, &styles.eur)?;
    }
    pen.text(TARGET_ROW, LABEL_COL, "Target", &styles.label_bold)?;
    pen.number(TARGET_ROW, AMOUNT_COL, TARGET, &styles.eur_bold)?;
    let sheet = pen.sheet();
    sheet.set_column_width(LABEL_COL, LABEL_WIDTH)?;
    sheet.set_column_width(AMOUNT_COL, AMOUNT_WIDTH)?;
    tally += pen.tally;
    Ok(tally)
}

/// A1 range of the amounts column: the Find a combination input.
pub fn amounts_address() -> String {
    range(FIRST_ROW, AMOUNT_COL, FIRST_ROW + COUNT - 1, AMOUNT_COL)
}

/// A1 address of the Target cell, for the checklist.
pub fn target_address() -> String {
    cell(TARGET_ROW, AMOUNT_COL)
}
