// Where things sit in the model sheets: the year columns, the Assumptions and
// P&L row maps and the A1 helpers. One source for formulas, charts, notes and
// the checklist, so a moved row changes in exactly one place.

pub const YEARS: [&str; 6] = ["2024A", "2025E", "2026E", "2027E", "2028E", "2029E"];
pub const LABEL_COL: u16 = 0;
pub const UNIT_COL: u16 = 1;
pub const FIRST_YEAR_COL: u16 = 2;
pub const TITLE_ROW: u32 = 0;
pub const HEADER_ROW: u32 = 2;
pub const PNL_SHEET: &str = "P&L";
pub const ASSUMPTIONS_SHEET: &str = "Assumptions";

/// Assumptions rows; the forecast years 2025E-2029E start in FIRST_YEAR_COL.
pub mod assumptions {
    pub const GROWTH: u32 = 3;
    pub const GROSS_MARGIN: u32 = 4;
    pub const OPEX_RATIO: u32 = 5;
    pub const DA_RATIO: u32 = 6;
    pub const INTEREST: u32 = 7;
    pub const TAX_RATE: u32 = 8;
    pub const REVENUE_2024: u32 = 10;
}

/// P&L rows; 2024A sits in FIRST_YEAR_COL, the forecast years follow.
pub mod pnl {
    pub const REVENUE: u32 = 3;
    pub const GROWTH: u32 = 4;
    pub const COGS: u32 = 5;
    pub const GROSS_PROFIT: u32 = 6;
    pub const GROSS_MARGIN: u32 = 7;
    pub const OPEX: u32 = 8;
    pub const EBITDA: u32 = 9;
    pub const EBITDA_MARGIN: u32 = 10;
    pub const DA: u32 = 11;
    pub const EBIT: u32 = 12;
    pub const INTEREST: u32 = 13;
    pub const EBT: u32 = 14;
    pub const TAX: u32 = 15;
    pub const NET_INCOME: u32 = 16;
    pub const NET_MARGIN: u32 = 17;
    pub const PER_MONTH: u32 = 18;
    pub const CHECK: u32 = 19;
}

/// Column of P&L year `i` (0 = 2024A).
pub const fn year_col(i: usize) -> u16 {
    FIRST_YEAR_COL + i as u16
}

/// Column of the Assumptions entry that drives P&L year `i` (2025E = first).
pub const fn assumption_col(i: usize) -> u16 {
    FIRST_YEAR_COL + i.saturating_sub(1) as u16
}

pub const fn last_year_col() -> u16 {
    FIRST_YEAR_COL + YEARS.len() as u16 - 1
}

/// 0-based column index to Excel letters: 0 -> A, 26 -> AA.
pub fn column(col: u16) -> String {
    let mut n = u32::from(col) + 1;
    let mut letters = Vec::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        letters.push(char::from(b'A' + rem as u8));
        n = (n - 1) / 26;
    }
    letters.iter().rev().collect()
}

pub fn a1_row(row: u32) -> String {
    (row + 1).to_string()
}

pub fn cell(row: u32, col: u16) -> String {
    format!("{}{}", column(col), a1_row(row))
}

pub fn cell_abs(row: u32, col: u16) -> String {
    format!("${}${}", column(col), a1_row(row))
}

/// A quoted reference into the P&L sheet, relative or absolute.
pub fn pnl_ref(row: u32, col: u16) -> String {
    format!("'{PNL_SHEET}'!{}", cell(row, col))
}

pub fn pnl_ref_abs(row: u32, col: u16) -> String {
    format!("'{PNL_SHEET}'!{}", cell_abs(row, col))
}

pub fn assumptions_ref_abs(row: u32, col: u16) -> String {
    format!("{ASSUMPTIONS_SHEET}!{}", cell_abs(row, col))
}

/// An A1 range from two corners.
pub fn range(first_row: u32, first_col: u16, last_row: u32, last_col: u16) -> String {
    format!("{}:{}", cell(first_row, first_col), cell(last_row, last_col))
}
