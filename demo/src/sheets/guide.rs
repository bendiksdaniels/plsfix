// The 7-row guide band every visible sheet opens with: what to try here, up
// to four numbered tasks, then the exact ribbon buttons and shortcut keys.
// Column A only; no wrap, so the line overflows across the empty cells to
// its right instead of breaking inside a narrow column.

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::{GUIDE_ROWS, LABEL_COL};
use crate::pen::Pen;
use crate::style::Styles;
use crate::tally::Tally;

/// One sheet's guide band content: a purpose line, up to four numbered
/// tasks, and the ribbon buttons and shortcut keys that do them.
pub struct Guide {
    pub title: &'static str,
    pub tasks: &'static [&'static str],
    pub commands: &'static [&'static str],
}

/// At most four tasks fit between the title and the Commands line.
pub(crate) const MAX_TASKS: usize = 4;
const TITLE_ROW: u32 = 0;
const FIRST_TASK_ROW: u32 = TITLE_ROW + 1;
const COMMANDS_ROW: u32 = FIRST_TASK_ROW + MAX_TASKS as u32;
const BLANK_ROW: u32 = COMMANDS_ROW + 1;

// The band's own geometry must add up to exactly GUIDE_ROWS, or the sheet
// content layout.rs shifts by GUIDE_ROWS would overlap the band or leave a
// gap above it.
const _: () = assert!(BLANK_ROW + 1 == GUIDE_ROWS);

/// Writes the band at the top of `sheet` (rows 0..GUIDE_ROWS) and returns
/// what it wrote, for the caller to fold into its own tally.
pub fn write(sheet: &mut Worksheet, styles: &Styles, guide: &Guide) -> Result<Tally, XlsxError> {
    // Each sheet file that builds a Guide at compile time also asserts its
    // own task count (a build error, not this); this is the backstop for
    // the one that cannot (pnl.rs, whose third task is computed at run
    // time), so a fifth task is never silently dropped there either.
    assert!(guide.tasks.len() <= MAX_TASKS, "{}: at most four tasks fit the guide band", guide.title);
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, &format!("Try on this sheet: {}", guide.title), &styles.guide_title)?;
    for (i, task) in guide.tasks.iter().enumerate() {
        let row = FIRST_TASK_ROW + i as u32;
        pen.text(row, LABEL_COL, &format!("{}. {task}", i + 1), &styles.guide_text)?;
    }
    let commands = format!("Commands: {}", guide.commands.join(", "));
    pen.text(COMMANDS_ROW, LABEL_COL, &commands, &styles.guide_text)?;
    // Paint the tint over the whole band, not just column A: every band
    // cell already carries it via its own format, so this only reaches the
    // empty cells to the right where the overflowing text visually sits.
    let sheet = pen.sheet();
    for row in 0..GUIDE_ROWS {
        sheet.set_row_format(row, &styles.guide_text)?;
    }
    Ok(pen.tally)
}
