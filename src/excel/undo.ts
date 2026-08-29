// pls,fix Undo: one slot, capturing what a mutating action is about to overwrite so
// it can be restored on demand. Office.js writes never reach Excel's own undo
// stack, so this is the pane's only safety net for a single last action.

import { SELECTION_CELL_CAP } from "./internal";
import { parseAddress } from "./shared";

// One captured rectangle. A ctrl-clicked selection is several of them, and an
// action that writes into every area has to be able to put every area back.
interface UndoBlock {
  sheetId: string;
  address: string;
  formulas: (string | number | boolean)[][];
  numberFormat: string[][];
  formats: Excel.CellProperties[][];
}

interface UndoSlot {
  label: string;
  blocks: UndoBlock[];
}

let undoSlot: UndoSlot | null = null;

// The full settable surface, so a restore is not partial: fills, fonts,
// borders, alignment, wrapping and indent all come back (row height cannot).
function requestFormats(
  range: Excel.Range,
): OfficeExtension.ClientResult<Excel.CellProperties[][]> {
  return range.getCellProperties({
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
}

// Office.js writes never reach Excel's own undo stack, so every mutating action
// stores what it is about to overwrite here first (user gap #5: undo trust).
export async function captureUndo(
  context: Excel.RequestContext,
  range: Excel.Range,
): Promise<void> {
  await captureUndoAreas(context, [range]);
}

/** The same capture over every area of a multi-area selection. */
export async function captureUndoAreas(
  context: Excel.RequestContext,
  ranges: Excel.Range[],
): Promise<void> {
  for (const range of ranges) range.load("address,rowCount,columnCount");
  await context.sync();

  // A skipped capture must not leave an older slot behind: the pane would then
  // offer to restore something that is not the last action.
  undoSlot = null;
  const cells = ranges.reduce(
    (total, range) => total + range.rowCount * range.columnCount,
    0,
  );
  if (cells > SELECTION_CELL_CAP) {
    undoSkipped = true;
    return;
  }
  undoSkipped = false;

  const pending = ranges.map((range) => {
    const properties = requestFormats(range);
    const sheet = range.worksheet;
    sheet.load("id");
    range.load("formulas,numberFormat");
    return { range, sheet, properties };
  });
  await context.sync();

  undoSlot = {
    label: pending.map(({ range }) => range.address).join(", "),
    blocks: pending.map(({ range, sheet, properties }) => ({
      sheetId: sheet.id,
      address: parseAddress(range.address).address,
      formulas: range.formulas as (string | number | boolean)[][],
      numberFormat: range.numberFormat as string[][],
      formats: properties.value,
    })),
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
    const sheets = slot.blocks.map((block) =>
      context.workbook.worksheets.getItemOrNullObject(block.sheetId),
    );
    for (const sheet of sheets) sheet.load("isNullObject");
    await context.sync();

    if (sheets.some((sheet) => sheet.isNullObject)) {
      undoSlot = null;
      throw new Error("The sheet that action ran on is gone.");
    }

    slot.blocks.forEach((block, index) => {
      const range = sheets[index]!.getRange(block.address);
      range.formulas = block.formulas;
      range.numberFormat = block.numberFormat;
      range.setCellProperties(
        block.formats as Excel.SettableCellProperties[][],
      );
    });
    await context.sync();

    // Only a restore that landed consumes the slot; a failed one stays retryable.
    undoSlot = null;
    return slot.label;
  });
}
