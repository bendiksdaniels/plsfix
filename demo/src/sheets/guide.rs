// The 7-row guide band every visible sheet opens with: what to try here, up
// to four numbered tasks, then the exact ribbon buttons and shortcut keys.
// Column A only (no merges, so range tools reading the sheet stay simple).

use rust_xlsxwriter::{Worksheet, XlsxError};

use crate::layout::LABEL_COL;
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

const TITLE_ROW: u32 = 0;
const FIRST_TASK_ROW: u32 = 1;
const MAX_TASKS: usize = 4;
const COMMANDS_ROW: u32 = 5;

/// Writes the band at the top of `sheet` (rows 0-6) and returns what it
/// wrote, for the caller to fold into its own tally.
pub fn write(sheet: &mut Worksheet, styles: &Styles, guide: &Guide) -> Result<Tally, XlsxError> {
    debug_assert!(guide.tasks.len() <= MAX_TASKS, "{}: at most four tasks fit the band", guide.title);
    let mut pen = Pen::new(sheet);
    pen.text(TITLE_ROW, LABEL_COL, &format!("Try on this sheet: {}", guide.title), &styles.guide_title)?;
    for (i, task) in guide.tasks.iter().take(MAX_TASKS).enumerate() {
        let row = FIRST_TASK_ROW + i as u32;
        pen.text(row, LABEL_COL, &format!("{}. {task}", i + 1), &styles.guide_text)?;
    }
    let commands = format!("Commands: {}", guide.commands.join(", "));
    pen.text(COMMANDS_ROW, LABEL_COL, &commands, &styles.guide_text)?;
    Ok(pen.tally)
}
