// The pen every sheet builder writes with: it forwards to the worksheet and
// counts what it wrote, so the tally and the file can never drift apart.
// Invariant: every cell a builder writes goes through one of the three methods.

use rust_xlsxwriter::{Format, Worksheet, XlsxError};

use crate::tally::Tally;

pub struct Pen<'a> {
    sheet: &'a mut Worksheet,
    pub tally: Tally,
}

impl<'a> Pen<'a> {
    pub fn new(sheet: &'a mut Worksheet) -> Pen<'a> {
        Pen { sheet, tally: Tally::default() }
    }

    /// The worksheet itself, for widths, panes, notes and charts (no cells).
    pub fn sheet(&mut self) -> &mut Worksheet {
        self.sheet
    }

    pub fn text(&mut self, row: u32, col: u16, text: &str, format: &Format) -> Result<&mut Self, XlsxError> {
        self.sheet.write_string_with_format(row, col, text, format)?;
        self.tally.strings += 1;
        Ok(self)
    }

    pub fn number(&mut self, row: u32, col: u16, value: f64, format: &Format) -> Result<&mut Self, XlsxError> {
        self.sheet.write_number_with_format(row, col, value, format)?;
        self.tally.numbers += 1;
        Ok(self)
    }

    pub fn formula(&mut self, row: u32, col: u16, formula: &str, format: &Format) -> Result<&mut Self, XlsxError> {
        self.sheet.write_formula_with_format(row, col, formula, format)?;
        self.tally.formulas += 1;
        Ok(self)
    }
}
