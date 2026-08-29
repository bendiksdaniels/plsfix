// pls,fix Undo: one slot, capturing what a mutating action is about to overwrite so
// it can be restored on demand. Office.js writes never reach Excel's own undo
// stack, so this is the pane's only safety net for a single last action.

import { SELECTION_CELL_CAP } from "./internal";
import { parseAddress } from "./shared";

interface UndoSlot {
  sheetId: string;
  address: string;
  label: string;
  formulas: (string | number | boolean)[][];
  numberFormat: string[][];
  formats: Excel.CellProperties[][];
}

let undoSlot: UndoSlot | null = null;

// Office.js writes never reach Excel's own undo stack, so every mutating action
// stores what it is about to overwrite here first (user gap #5: undo trust).
export async function captureUndo(
  context: Excel.RequestContext,
  range: Excel.Range,
): Promise<void> {
  range.load("address,rowCount,columnCount");
  await context.sync();

  // A skipped capture must not leave an older slot behind: the pane would then
  // offer to restore something that is not the last action.
  undoSlot = null;
  if (range.rowCount * range.columnCount > SELECTION_CELL_CAP) {
    undoSkipped = true;
    return;
  }
  undoSkipped = false;

  // The full settable surface, so a restore is not partial: fills, fonts,
  // borders, alignment, wrapping and indent all come back (row height cannot).
  const properties = range.getCellProperties({
    format: {
      fill: { color: true, pattern: true, patternColor: true },
      font: {
        bold: true,
        color: true,
        italic: true,
        name: true,
        size: true,
        underline: true,
      },
      borders: { color: true, style: true, weight: true },
      horizontalAlignment: true,
      verticalAlignment: true,
      wrapText: true,
      indentLevel: true,
    },
  });
  const sheet = range.worksheet;
  sheet.load("id");
  range.load("formulas,numberFormat");
  await context.sync();

  undoSlot = {
    sheetId: sheet.id,
    address: parseAddress(range.address).address,
    label: range.address,
    formulas: range.formulas as (string | number | boolean)[][],
    numberFormat: range.numberFormat as string[][],
    formats: properties.value,
  };
}

export function undoTarget(): string | null {
  return undoSlot?.label ?? null;
}

let undoSkipped = false;

// True when the most recent mutating action ran without undo protection
// (selection over the cap) — the pane says so instead of implying a safety net.
// Read-once: a later non-mutating action must not inherit the flag.
export function lastUndoSkipped(): boolean {
  const was = undoSkipped;
  undoSkipped = false;
  return was;
}

export async function undoLastAction(): Promise<string> {
  const slot = undoSlot;
  if (!slot) throw new Error("There is no pls,fix action to undo yet.");

  return Excel.run(async (context) => {
    // Sheet id rather than name, so a rename between action and undo is fine.
    const sheet = context.workbook.worksheets.getItemOrNullObject(slot.sheetId);
    sheet.load("isNullObject");
    await context.sync();

    if (sheet.isNullObject) {
      undoSlot = null;
      throw new Error("The sheet that action ran on is gone.");
    }

    const range = sheet.getRange(slot.address);
    range.formulas = slot.formulas;
    range.numberFormat = slot.numberFormat;
    range.setCellProperties(slot.formats as Excel.SettableCellProperties[][]);
    await context.sync();

    // Only a restore that landed consumes the slot; a failed one stays retryable.
    undoSlot = null;
    return slot.label;
  });
}
