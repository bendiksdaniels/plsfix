// The "Start here" sheet: the play-around checklist, one row per tool with
// where to click, what to do and what should happen. Every address comes from
// the layout or the sheet that owns it, so a moved block cannot go stale here.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{
    a1_row, assumptions as a, cell, last_year_col, pnl, range, year_col, ASSUMPTIONS_SHEET,
    FIRST_YEAR_COL, GUIDE_ROWS, HEADER_ROW, LABEL_COL, PNL_SHEET, TITLE_ROW,
};
use crate::pen::Pen;
use crate::sheets::guide::{self, Guide};
use crate::sheets::pnl::planted_cells;
use crate::sheets::{bridge, data, rounding, scratch, sensitivity, variance};
use crate::style::Styles;
use crate::tally::Tally;

pub const NAME: &str = "Start here";
const NOTE_ROW: u32 = GUIDE_ROWS + 1;
const TABLE_HEADER_ROW: u32 = GUIDE_ROWS + 3;
const FIRST_STEP_ROW: u32 = GUIDE_ROWS + 4;
const GUIDE: Guide = Guide {
    title: "orientation and the full checklist",
    tasks: &[
        "Open the pane (Ctrl+Shift+M).",
        "Walk the checklist below, one row per tool.",
        "Open the demo deck \"pls,fix Demo Deck.pptx\" for the PowerPoint tasks.",
    ],
    commands: &["Open pls,fix (Ctrl+Shift+M)"],
};
const NUMBER_COL: u16 = 0;
const TOOL_COL: u16 = 1;
const PLACE_COL: u16 = 2;
const ACTION_COL: u16 = 3;
const OUTCOME_COL: u16 = 4;
const COLUMNS: [(&str, f64); 5] = [
    ("#", 4.0),
    ("Tool", 30.0),
    ("Where", 28.0),
    ("Do this", 92.0),
    ("You should see", 78.0),
];
const HOST: &str = "dbautomatizacijas.com/modelis";

struct Step {
    tool: &'static str,
    place: String,
    action: String,
    outcome: String,
}

fn step(tool: &'static str, place: impl Into<String>, action: impl Into<String>, outcome: impl Into<String>) -> Step {
    Step { tool, place: place.into(), action: action.into(), outcome: outcome.into() }
}

pub fn build(sheet: &mut Worksheet, styles: &Styles) -> Result<Tally, XlsxError> {
    let mut tally = guide::write(sheet, styles, &GUIDE)?;
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, "pls,fix: play-around workbook", &styles.title)?;
    pen.text(NOTE_ROW, LABEL_COL, &format!("Ribbon tab \"pls,fix\" in Excel and in PowerPoint; the pane loads from {HOST}. DemoCo SIA is invented."), &styles.note)?;
    for (i, (title, _)) in COLUMNS.iter().enumerate() {
        pen.text(TABLE_HEADER_ROW, i as u16, title, &styles.header_left)?;
    }
    for (i, step) in steps().iter().enumerate() {
        let row = FIRST_STEP_ROW + i as u32;
        pen.number(row, NUMBER_COL, (i + 1) as f64, &styles.whole)?;
        pen.text(row, TOOL_COL, step.tool, &styles.label_bold)?;
        pen.text(row, PLACE_COL, &step.place, &styles.label)?;
        pen.text(row, ACTION_COL, &step.action, &styles.label)?;
        pen.text(row, OUTCOME_COL, &step.outcome, &styles.label)?;
    }
    let sheet = pen.sheet();
    for (i, (_, width)) in COLUMNS.iter().enumerate() {
        sheet.set_column_width(i as u16, *width)?;
    }
    sheet.set_active(true);
    tally += pen.tally;
    Ok(tally)
}

fn steps() -> Vec<Step> {
    let mut steps = model_steps();
    steps.extend(chart_steps());
    steps.extend(workbook_steps());
    steps.extend(link_steps());
    steps
}

fn model_steps() -> Vec<Step> {
    let last = last_year_col();
    let planted = planted_cells();
    vec![
        step("Open the pane", "Excel ribbon", "Click the pls,fix tab, then Model Tools; the pane has the tabs Tools, Workbook, Links and Brand", format!("The pane opens on the right, loaded from {HOST}")),
        step("Autocolor selection", PNL_SHEET, format!("Select {} and press Autocolor selection (Tools tab)", range(HEADER_ROW, LABEL_COL, pnl::CHECK, last)), format!("Inputs blue, formulas black, cross-sheet links green; {} and {} stand out", planted[0], planted[1])),
        step("Audit overlay", PNL_SHEET, format!("Select {} and press Audit overlay (press again to clear)", range(pnl::REVENUE, FIRST_YEAR_COL, pnl::NET_MARGIN, last)), format!("Inconsistent formulas striped: {} carries a hardcoded 0.21", planted[0])),
        step("Precedents / Dependents", PNL_SHEET, format!("Click {} (EBITDA 2025E) and press Precedents; on {}!{} press Dependents", cell(pnl::EBITDA, year_col(1)), ASSUMPTIONS_SHEET, cell(a::GROWTH, FIRST_YEAR_COL)), "The cells it reads, then the cells that read it, get selected"),
        step("Fill formula right", PNL_SHEET, format!("Select {} and press Fill formula right", range(pnl::PER_MONTH, FIRST_YEAR_COL, pnl::PER_MONTH, last)), "The first cell's formula is filled across the row"),
        step("Number formats, x1000, /1000, Sign flip", PNL_SHEET, "Select some numbers and press the buttons on the Tools tab; Undo last pls,fix action takes the last one back", "Formats cycle; values scale or flip sign"),
        step("CAGR", PNL_SHEET, format!("Select {} (Revenue) and press CAGR", range(pnl::REVENUE, FIRST_YEAR_COL, pnl::REVENUE, last)), "The 2024A-2029E CAGR written beside the row"),
        step("Find a combination", variance::NAME, format!("Select {} and type the number from {} into the pane's Find a combination, then press Find cells", variance::amounts_address(), variance::target_address()), "The invoice lines that add up to the target become the selection; nothing else in the list matches"),
    ]
}

fn chart_steps() -> Vec<Step> {
    vec![
        step("Waterfall from selection", bridge::NAME, format!("Select {} and press Waterfall from selection", bridge::table_address()), "A native bridge chart: opening and closing totals, the steps between"),
        step("CAGR label / Brand-format chart", bridge::NAME, "Click the Revenue chart, press CAGR label, then Brand-format chart", "A CAGR label on the chart; brand formatting applied (the pie on Rounding is there to export too)"),
        step("Tornado from selection", sensitivity::NAME, format!("Select {} and press Tornado from selection", sensitivity::block_address()), format!("A ranked sensitivity chart around the base case in {}", sensitivity::base_address())),
        step("Consistent rounding / =PLSFIX.ROUND", rounding::NAME, format!("Select {} and press Consistent rounding, or type the formulas shown in the last column", rounding::points_address()), "PLSFIX.ROUND formulas beside the selection whose parts add up to the rounded total (the Excel ROUND column sums to 101)"),
        step("Unpivot selection", data::NAME, format!("Select {} and press Unpivot selection", data::grid_address()), "A new sheet with one Row / Column / Value line per cell of the grid"),
    ]
}

fn workbook_steps() -> Vec<Step> {
    vec![
        step("Find in workbook", "Workbook tab", "Search for \"export\"", format!("Hits on {}: a text cell and a note", data::NAME)),
        step("Insert contents sheet, Scan broken names, Scan styles, Prepare for sharing", "Workbook tab", "Press each button", format!("A contents sheet; {} reported as broken; every sheet back to A1 and the hidden {} sheet reported", crate::sheets::BROKEN_NAME, scratch::NAME)),
    ]
}

fn link_steps() -> Vec<Step> {
    let last = last_year_col();
    vec![
        step("Export to PowerPoint", "Links tab", format!("Select {} on {} and press Export selection; click the Revenue chart on {} and press Export active chart", range(HEADER_ROW, LABEL_COL, pnl::NET_INCOME, last), PNL_SHEET, bridge::NAME), "Two links in the list; both wait in the PowerPoint Inbox (if it says unpaired: Generate and Copy the key on this tab, paste it in PowerPoint Settings, Save key)"),
        step("Paste latest linked", "PowerPoint: pls,fix tab, Inbox", "With both exports waiting, click Paste latest linked", "The most recent export (the Revenue chart) lands on the active slide in one click, no need to open the list"),
        step("Insert and update", "PowerPoint: pls,fix tab, Links", format!("Insert the remaining item (the P&L table) from the Inbox, then move and resize both objects; in Excel set {}!{} to 12% and press Push all; in PowerPoint press Update all", ASSUMPTIONS_SHEET, cell(a::GROWTH, FIRST_YEAR_COL)), "Both objects refresh in place with the new numbers, sizes kept"),
        step("Links list filters", "PowerPoint: pls,fix tab, Links", "Type part of a slide, object or workbook name in the search box, then try the Source workbook, Slide and Link status dropdowns", "The link table narrows to matching rows; a ticked row stays ticked even if the filter hides it"),
        step("Source missing", PNL_SHEET, format!("Delete rows {}-{} (the whole exported block) and press Push all; then undo with Cmd+Z and push again", a1_row(HEADER_ROW), a1_row(pnl::NET_INCOME)), "Push reports 1 missing and the link reads \"source missing\"; after the undo it pushes again (deleting rows inside the block only shrinks it)"),
        step("Align, distribute, match size, select similar, swap, Smart Painter", "PowerPoint: pls,fix tab, Tools", "Insert 3-4 shapes on a slide, select them and try Align, Distribute, Match size, Select similar and Swap; then Capture one shape's style under Smart Painter and Apply it to another", "Shapes line up or space out, match the first one's size, swap places, or pick up its fill and outline"),
    ]
}
