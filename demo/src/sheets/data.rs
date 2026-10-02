// A wide product-by-month grid for Unpivot, with a note and a text cell that
// Super Find can hit.

use rust_xlsxwriter::{Note, Worksheet, XlsxError};

use crate::layout::{cell, GUIDE_ROWS, LABEL_COL, TITLE_ROW};
use crate::pen::Pen;
use crate::sheets::guide::{self, Guide};
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = "Data";
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const HEADER_ROW: u32 = GUIDE_ROWS + 2;
const FIRST_PRODUCT_ROW: u32 = GUIDE_ROWS + 3;
const PRODUCT_COL: u16 = 0;
const FIRST_MONTH_COL: u16 = 1;
const HINT_ROW: u32 = GUIDE_ROWS + 9;
const FIND_ROW: u32 = GUIDE_ROWS + 10;
const GUIDE: Guide = Guide {
    title: "a wide product-by-month grid",
    tasks: &[
        "Unpivot the product-by-month grid into one row per cell.",
        "Find in workbook searches for the note's text.",
        "Export selection as a picture, then in PowerPoint Slide 4: select the empty placeholder, Where = Selected shape.",
    ],
    commands: &["Find in workbook (Ctrl+Shift+Alt+F)"],
};
const _: () = assert!(GUIDE.tasks.len() <= guide::MAX_TASKS, "a fifth task would land on the Commands row");
const NOTE_MONTH: usize = 2;
const NOTE_PRODUCT: usize = 2;
const PRODUCT_WIDTH: f64 = 16.0;
const MONTH_WIDTH: f64 = 9.0;

const MONTHS: [&str; 6] = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
const PRODUCTS: [(&str, [f64; 6]); 5] = [
    ("Rye bread", [412.0, 398.0, 455.0, 430.0, 441.0, 468.0]),
    ("Wheat bread", [388.0, 375.0, 402.0, 396.0, 410.0, 425.0]),
    ("Pastries", [265.0, 258.0, 301.0, 288.0, 295.0, 310.0]),
    ("Crispbread", [142.0, 150.0, 188.0, 160.0, 158.0, 171.0]),
    ("Cakes", [96.0, 101.0, 120.0, 118.0, 125.0, 133.0]),
];

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "Monthly sales by product (EUR thousands): a wide grid", &styles.title)?;
    pen.text(NOTE_ROW, LABEL_COL, &format!("Select {} and press Unpivot for the long product-month-value table.", grid_address()), &styles.note)?;
    pen.text(HEADER_ROW, PRODUCT_COL, "Product", &styles.header_left)?;
    for (i, month) in MONTHS.iter().enumerate() {
        pen.text(HEADER_ROW, FIRST_MONTH_COL + i as u16, month, &styles.header)?;
    }
    for (i, (product, values)) in PRODUCTS.iter().enumerate() {
        let row = FIRST_PRODUCT_ROW + i as u32;
        pen.text(row, PRODUCT_COL, product, &styles.label)?;
        for (m, value) in values.iter().enumerate() {
            pen.number(row, FIRST_MONTH_COL + m as u16, *value, &styles.eur)?;
        }
    }
    pen.text(HINT_ROW, LABEL_COL, "Super Find: search for \"export\" on the Workbook tab; it is in the text below and in the note on the March pastries cell.", &styles.note)?;
    pen.text(FIND_ROW, LABEL_COL, "Export order to the Riga terminal booked for March.", &styles.label)?;
    let sheet = pen.sheet();
    sheet.insert_note(
        FIRST_PRODUCT_ROW + NOTE_PRODUCT as u32,
        FIRST_MONTH_COL + NOTE_MONTH as u16,
        &Note::new("Check with the sales team: March includes a one-off export order."),
    )?;
    sheet.set_column_width(PRODUCT_COL, PRODUCT_WIDTH)?;
    sheet.set_column_range_width(FIRST_MONTH_COL, FIRST_MONTH_COL + MONTHS.len() as u16 - 1, MONTH_WIDTH)?;
    tally += pen.tally;
    Ok(tally)
}

/// A1 address of the wide grid, header row and product column included.
pub fn grid_address() -> String {
    format!(
        "{}:{}",
        cell(HEADER_ROW, PRODUCT_COL),
        cell(FIRST_PRODUCT_ROW + PRODUCTS.len() as u32 - 1, FIRST_MONTH_COL + MONTHS.len() as u16 - 1)
    )
}
