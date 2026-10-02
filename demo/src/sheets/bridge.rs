// The EBITDA bridge 2024A -> 2025E as a two-column table (the shape the
// Waterfall button wants: labels left, values right, totals first and last),
// plus two ready charts for the chart tools and for "Export chart".

use rust_xlsxwriter::{
    Chart, ChartDataLabel, ChartDataLabelPosition, ChartLine, ChartSolidFill, ChartType,
    Worksheet, XlsxError,
};

use crate::layout::{
    cell, last_year_col, pnl::*, pnl_ref, year_col, FIRST_YEAR_COL, GUIDE_ROWS, HEADER_ROW,
    LABEL_COL, PNL_SHEET, TITLE_ROW,
};
use crate::pen::Pen;
use crate::sheets::guide::{self, Guide};
use crate::style::{Styles, TEAL, NAVY, EUR_K};
use crate::tally::Tally;

pub const NAME: &str = "Bridge";
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const STEP_COL: u16 = 0;
const VALUE_COL: u16 = 1;
const FIRST_STEP_ROW: u32 = GUIDE_ROWS + 3;
const STEP_COUNT: u32 = 5;
const CHECK_ROW: u32 = GUIDE_ROWS + 9;
const HINT_ROW: u32 = GUIDE_ROWS + 11;
const CHART_COL: u16 = 3;
const REVENUE_CHART_ROW: u32 = GUIDE_ROWS + 2;
const MARGIN_CHART_ROW: u32 = GUIDE_ROWS + 19;
const GUIDE: Guide = Guide {
    title: "the EBITDA bridge, table and charts",
    tasks: &[
        "Press Waterfall from selection on the two-column table.",
        "Export active chart (the Revenue chart), then in PowerPoint Slide 3, Where = Left half.",
    ],
    commands: &["Waterfall from selection (Ctrl+Shift+B)"],
};
const _: () = assert!(GUIDE.tasks.len() <= guide::MAX_TASKS, "a fifth task would land on the Commands row");
const CHART_WIDTH: u32 = 560;
const CHART_HEIGHT: u32 = 300;
const BAR_GAP: u16 = 60;
const LINE_WIDTH: f64 = 2.25;
const STEP_WIDTH: f64 = 24.0;
const VALUE_WIDTH: f64 = 11.0;

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "EBITDA bridge 2024A to 2025E (EUR thousands)", &styles.title)?;
    pen.text(NOTE_ROW, LABEL_COL, "Select the two columns below and press Waterfall; click a chart and use Chart CAGR or Export chart.", &styles.note)?;
    pen.text(HEADER_ROW, STEP_COL, "Step", &styles.header_left)?;
    pen.text(HEADER_ROW, VALUE_COL, "EUR k", &styles.header)?;
    for (i, (label, formula)) in steps().iter().enumerate() {
        let row = FIRST_STEP_ROW + i as u32;
        let total = i == 0 || i as u32 == STEP_COUNT - 1;
        let label_format = if total { &styles.label_bold } else { &styles.label };
        let value_format = if total { &styles.eur_bold } else { &styles.eur };
        pen.text(row, STEP_COL, label, label_format)?;
        pen.formula(row, VALUE_COL, formula, value_format)?;
    }
    pen.text(CHECK_ROW, STEP_COL, "Check: opening + steps - closing (must be 0)", &styles.label)?;
    pen.formula(CHECK_ROW, VALUE_COL, &check_formula(), &styles.eur)?;
    pen.text(HINT_ROW, STEP_COL, &format!("Waterfall input: {}", table_address()), &styles.note)?;
    let sheet = pen.sheet();
    sheet.set_column_width(STEP_COL, STEP_WIDTH)?;
    sheet.set_column_width(VALUE_COL, VALUE_WIDTH)?;
    sheet.insert_chart(REVENUE_CHART_ROW, CHART_COL, &revenue_chart())?;
    sheet.insert_chart(MARGIN_CHART_ROW, CHART_COL, &margin_chart())?;
    tally += pen.tally;
    Ok(tally)
}

/// A1 address of the bridge table, for the checklist and the hint.
pub fn table_address() -> String {
    format!(
        "{}:{}",
        cell(FIRST_STEP_ROW, STEP_COL),
        cell(FIRST_STEP_ROW + STEP_COUNT - 1, VALUE_COL)
    )
}

fn steps() -> [(&'static str, String); STEP_COUNT as usize] {
    let (c24, c25) = (year_col(0), year_col(1));
    [
        ("EBITDA 2024A", format!("={}", pnl_ref(EBITDA, c24))),
        (
            "Revenue growth",
            format!("=({}-{})*{}", pnl_ref(REVENUE, c25), pnl_ref(REVENUE, c24), pnl_ref(GROSS_MARGIN, c24)),
        ),
        (
            "Gross margin",
            format!("=({}-{})*{}", pnl_ref(GROSS_MARGIN, c25), pnl_ref(GROSS_MARGIN, c24), pnl_ref(REVENUE, c25)),
        ),
        ("Operating expenses", format!("={}-{}", pnl_ref(OPEX, c25), pnl_ref(OPEX, c24))),
        ("EBITDA 2025E", format!("={}", pnl_ref(EBITDA, c25))),
    ]
}

fn check_formula() -> String {
    let last_step = FIRST_STEP_ROW + STEP_COUNT - 2;
    let closing = FIRST_STEP_ROW + STEP_COUNT - 1;
    format!(
        "=SUM({}:{})-{}",
        cell(FIRST_STEP_ROW, VALUE_COL),
        cell(last_step, VALUE_COL),
        cell(closing, VALUE_COL)
    )
}

fn revenue_chart() -> Chart {
    let mut chart = Chart::new(ChartType::Column);
    chart.set_name("Revenue chart");
    chart.title().set_name("Revenue, EUR thousands");
    chart.legend().set_hidden();
    chart
        .add_series()
        .set_categories((PNL_SHEET, HEADER_ROW, FIRST_YEAR_COL, HEADER_ROW, last_year_col()))
        .set_values((PNL_SHEET, REVENUE, FIRST_YEAR_COL, REVENUE, last_year_col()))
        .set_format(ChartSolidFill::new().set_color(TEAL))
        .set_gap(BAR_GAP)
        .set_data_label(
            ChartDataLabel::new()
                .show_value()
                .set_position(ChartDataLabelPosition::OutsideEnd),
        );
    chart.y_axis().set_num_format(EUR_K);
    chart.set_width(CHART_WIDTH).set_height(CHART_HEIGHT);
    chart
}

fn margin_chart() -> Chart {
    let mut chart = Chart::new(ChartType::Line);
    chart.set_name("EBITDA margin chart");
    chart.title().set_name("EBITDA margin");
    chart.legend().set_hidden();
    chart
        .add_series()
        .set_categories((PNL_SHEET, HEADER_ROW, FIRST_YEAR_COL, HEADER_ROW, last_year_col()))
        .set_values((PNL_SHEET, EBITDA_MARGIN, FIRST_YEAR_COL, EBITDA_MARGIN, last_year_col()))
        .set_format(ChartLine::new().set_color(NAVY).set_width(LINE_WIDTH))
        .set_data_label(
            ChartDataLabel::new()
                .show_value()
                .set_position(ChartDataLabelPosition::Above),
        );
    chart.y_axis().set_num_format("0%");
    chart.set_width(CHART_WIDTH).set_height(CHART_HEIGHT);
    chart
}
