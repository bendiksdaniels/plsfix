// The P&L: DemoCo's 2024A actuals and a 2025E-2029E forecast driven from the
// Assumptions sheet, written from one line table with column and row
// placeholders. Two errors are planted on purpose (a hardcode inside a formula,
// a typed number in a formula row) for Autocolor and the audit overlay to find.

use rust_xlsxwriter::{Format, Worksheet, XlsxError};

use crate::layout::{
    a1_row, assumption_col, assumptions as a, cell, column, last_year_col, pnl::*, year_col,
    FIRST_YEAR_COL, HEADER_ROW, LABEL_COL, TITLE_ROW, UNIT_COL, YEARS,
};
use crate::pen::Pen;
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = crate::layout::PNL_SHEET;
const NOTE_ROW: u32 = 1;
const LABEL_WIDTH: f64 = 46.0;
const UNIT_WIDTH: f64 = 7.0;
const YEAR_WIDTH: f64 = 11.0;
const MONTHS: &str = "12";

#[derive(Clone, Copy)]
enum Actual {
    /// 2024A uses the forecast template too.
    Same,
    Number(f64),
    Formula(&'static str),
    Blank,
}

struct Line {
    row: u32,
    label: &'static str,
    unit: &'static str,
    actual: Actual,
    forecast: Option<&'static str>,
    percent: bool,
    bold: bool,
}

// Placeholders: {c} this year's column, {p} the previous one, {a} the
// Assumptions column for this forecast year, {a0} the first Assumptions
// column; {rev} .. {ni} P&L rows, {arev} .. {atax} Assumptions rows.
const LINES: [Line; 17] = [
    Line { row: REVENUE, label: "Revenue", unit: "EUR k", actual: Actual::Formula("=Assumptions!{a0}{arev}"), forecast: Some("={p}{rev}*(1+Assumptions!{a}{ag})"), percent: false, bold: true },
    Line { row: GROWTH, label: "  growth", unit: "%", actual: Actual::Blank, forecast: Some("={c}{rev}/{p}{rev}-1"), percent: true, bold: false },
    Line { row: COGS, label: "Cost of goods sold", unit: "EUR k", actual: Actual::Number(-7_688.0), forecast: Some("=-{c}{rev}*(1-Assumptions!{a}{agm})"), percent: false, bold: false },
    Line { row: GROSS_PROFIT, label: "Gross profit", unit: "EUR k", actual: Actual::Same, forecast: Some("={c}{rev}+{c}{cogs}"), percent: false, bold: true },
    Line { row: GROSS_MARGIN, label: "  gross margin", unit: "%", actual: Actual::Same, forecast: Some("={c}{gp}/{c}{rev}"), percent: true, bold: false },
    Line { row: OPEX, label: "Operating expenses", unit: "EUR k", actual: Actual::Number(-2_852.0), forecast: Some("=-{c}{rev}*Assumptions!{a}{aopex}"), percent: false, bold: false },
    Line { row: EBITDA, label: "EBITDA", unit: "EUR k", actual: Actual::Same, forecast: Some("={c}{gp}+{c}{opex}"), percent: false, bold: true },
    Line { row: EBITDA_MARGIN, label: "  EBITDA margin", unit: "%", actual: Actual::Same, forecast: Some("={c}{ebitda}/{c}{rev}"), percent: true, bold: false },
    Line { row: DA, label: "Depreciation and amortisation", unit: "EUR k", actual: Actual::Number(-434.0), forecast: Some("=-{c}{rev}*Assumptions!{a}{ada}"), percent: false, bold: false },
    Line { row: EBIT, label: "EBIT", unit: "EUR k", actual: Actual::Same, forecast: Some("={c}{ebitda}+{c}{da}"), percent: false, bold: true },
    Line { row: INTEREST, label: "Interest expense", unit: "EUR k", actual: Actual::Number(-190.0), forecast: Some("=-Assumptions!{a}{aint}"), percent: false, bold: false },
    Line { row: EBT, label: "Profit before tax", unit: "EUR k", actual: Actual::Same, forecast: Some("={c}{ebit}+{c}{int}"), percent: false, bold: false },
    Line { row: TAX, label: "Income tax", unit: "EUR k", actual: Actual::Formula("=-MAX(0,{c}{ebt})*0.2"), forecast: Some("=-MAX(0,{c}{ebt})*Assumptions!{a}{atax}"), percent: false, bold: false },
    Line { row: NET_INCOME, label: "Net income", unit: "EUR k", actual: Actual::Same, forecast: Some("={c}{ebt}+{c}{tax}"), percent: false, bold: true },
    Line { row: NET_MARGIN, label: "  net margin", unit: "%", actual: Actual::Same, forecast: Some("={c}{ni}/{c}{rev}"), percent: true, bold: false },
    Line { row: CHECK, label: "Check: EBITDA + D&A - EBIT (must be 0)", unit: "EUR k", actual: Actual::Same, forecast: Some("={c}{ebitda}+{c}{da}-{c}{ebit}"), percent: false, bold: false },
    Line { row: PER_MONTH, label: "Revenue per month (select the row, press Fill Right)", unit: "EUR k", actual: Actual::Formula("={c}{rev}/12"), forecast: None, percent: false, bold: false },
];

#[derive(Clone, Copy)]
enum Plant {
    Formula(&'static str),
    Number(f64),
}

/// The two planted errors: (row, column, what lands there instead).
const PLANTED: [(u32, u16, Plant); 2] = [
    (OPEX, year_col(3), Plant::Formula("=-{c}{rev}*0.21")),
    (DA, year_col(4), Plant::Number(-560.0)),
];

pub fn planted_cells() -> Vec<String> {
    PLANTED.iter().map(|(row, col, _)| cell(*row, *col)).collect()
}

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut pen = Pen::new(sheet);
    let planted = planted_cells();
    let note = format!(
        "Two errors are planted: a hardcode inside {} and a typed number in {}. Autocolor and the Audit overlay find them.",
        planted[0], planted[1]
    );
    pen.text(TITLE_ROW, LABEL_COL, "DemoCo SIA: profit and loss (EUR thousands)", &styles.title)?;
    pen.text(NOTE_ROW, LABEL_COL, &note, &styles.note)?;
    write_header(&mut pen, styles)?;
    for line in &LINES {
        write_line(&mut pen, styles, line)?;
    }
    let sheet = pen.sheet();
    sheet.set_column_width(LABEL_COL, LABEL_WIDTH)?;
    sheet.set_column_width(UNIT_COL, UNIT_WIDTH)?;
    sheet.set_column_range_width(FIRST_YEAR_COL, last_year_col(), YEAR_WIDTH)?;
    sheet.set_freeze_panes(HEADER_ROW + 1, FIRST_YEAR_COL)?;
    Ok(pen.tally)
}

fn write_header(pen: &mut Pen, styles: &Styles) -> Result<(), XlsxError> {
    pen.text(HEADER_ROW, LABEL_COL, "EUR thousands", &styles.header_left)?;
    pen.text(HEADER_ROW, UNIT_COL, "Unit", &styles.header_left)?;
    for (i, year) in YEARS.iter().enumerate() {
        pen.text(HEADER_ROW, year_col(i), year, &styles.header)?;
    }
    Ok(())
}

enum Spec {
    Number(f64),
    Formula(&'static str),
    Blank,
}

fn spec(line: &Line, year: usize) -> Spec {
    if year > 0 {
        return line.forecast.map_or(Spec::Blank, Spec::Formula);
    }
    match line.actual {
        Actual::Same => line.forecast.map_or(Spec::Blank, Spec::Formula),
        Actual::Number(value) => Spec::Number(value),
        Actual::Formula(template) => Spec::Formula(template),
        Actual::Blank => Spec::Blank,
    }
}

fn planted(row: u32, col: u16) -> Option<Plant> {
    PLANTED.iter().find(|(r, c, _)| *r == row && *c == col).map(|(_, _, plant)| *plant)
}

fn number_format<'s>(styles: &'s Styles, line: &Line) -> &'s Format {
    if line.percent {
        &styles.pct
    } else if line.bold {
        &styles.eur_bold
    } else {
        &styles.eur
    }
}

fn write_line(pen: &mut Pen, styles: &Styles, line: &Line) -> Result<(), XlsxError> {
    let label = if line.bold { &styles.label_bold } else { &styles.label };
    pen.text(line.row, LABEL_COL, line.label, label)?;
    pen.text(line.row, UNIT_COL, line.unit, &styles.label)?;
    let format = number_format(styles, line);
    for year in 0..YEARS.len() {
        let col = year_col(year);
        let spec = match planted(line.row, col) {
            Some(Plant::Formula(template)) => Spec::Formula(template),
            Some(Plant::Number(value)) => Spec::Number(value),
            None => spec(line, year),
        };
        match spec {
            Spec::Number(value) => pen.number(line.row, col, value, format)?,
            Spec::Formula(template) => pen.formula(line.row, col, &expand(template, year), format)?,
            Spec::Blank => continue,
        };
    }
    Ok(())
}

fn expand(template: &str, year: usize) -> String {
    let pairs = [
        ("{c}", column(year_col(year))),
        ("{p}", column(year_col(year.saturating_sub(1)))),
        ("{a}", column(assumption_col(year))),
        ("{a0}", column(FIRST_YEAR_COL)),
        ("{rev}", a1_row(REVENUE)),
        ("{cogs}", a1_row(COGS)),
        ("{gp}", a1_row(GROSS_PROFIT)),
        ("{opex}", a1_row(OPEX)),
        ("{ebitda}", a1_row(EBITDA)),
        ("{da}", a1_row(DA)),
        ("{ebit}", a1_row(EBIT)),
        ("{int}", a1_row(INTEREST)),
        ("{ebt}", a1_row(EBT)),
        ("{tax}", a1_row(TAX)),
        ("{ni}", a1_row(NET_INCOME)),
        ("{arev}", a1_row(a::REVENUE_2024)),
        ("{ag}", a1_row(a::GROWTH)),
        ("{agm}", a1_row(a::GROSS_MARGIN)),
        ("{aopex}", a1_row(a::OPEX_RATIO)),
        ("{ada}", a1_row(a::DA_RATIO)),
        ("{aint}", a1_row(a::INTEREST)),
        ("{atax}", a1_row(a::TAX_RATE)),
    ];
    let expanded = pairs
        .iter()
        .fold(template.to_string(), |acc, (key, value)| acc.replace(key, value));
    debug_assert!(!expanded.contains('{'), "unexpanded placeholder in {template}");
    let _ = MONTHS;
    expanded
}
