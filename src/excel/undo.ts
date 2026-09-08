// pls,fix Undo: a five-deep stack, newest first, capturing what a mutating
// action is about to overwrite so it can be restored on demand. Office.js
// writes never reach Excel's own undo stack, so this is the pane's only
// safety net for its last few actions.

import { SELECTION_CELL_CAP } from "./internal";
import { syncWrite } from "./protection";
import { parseAddress } from "./shared";
import { pushCapped } from "./undo-stack";

// One captured rectangle. A ctrl-clicked selection is several of them, and an
// action that writes into every area has to be able to put every area back.
interface UndoBlock {
  sheetId: string;
  address: string;
  formulas: (string | number | boolean)[][];
  numberFormat: string[][];
  formats: Excel.CellProperties[][];
}

interface UndoEntry {
  label: string;
  blocks: UndoBlock[];
  cells: number;
}

export const UNDO_DEPTH = 5;
export const UNDO_CELL_BUDGET = 25_000;

// Newest first: index 0 is what "Undo" acts on next.
let undoStack: UndoEntry[] = [];

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

  const cells = ranges.reduce(
    (total, range) => total + range.rowCount * range.columnCount,
    0,
  );
  if (cells > SELECTION_CELL_CAP) {
    // An uncaptured action may have touched anything, including ranges an
    // older entry still thinks it can restore, so the whole stack goes with
    // it rather than offering to restore something that is not safe to.
    undoStack = [];
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

  const entry: UndoEntry = {
    label: pending.map(({ range }) => range.address).join(", "),
    blocks: pending.map(({ range, sheet, properties }) => ({
      sheetId: sheet.id,
      address: parseAddress(range.address).address,
      formulas: range.formulas as (string | number | boolean)[][],
      numberFormat: range.numberFormat as string[][],
      formats: properties.value,
    })),
    cells,
  };
  undoStack = pushCapped(undoStack, entry, {
    maxDepth: UNDO_DEPTH,
    maxCells: UNDO_CELL_BUDGET,
  });
}

export function undoTarget(): string | null {
  return undoStack[0]?.label ?? null;
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
  const top = undoStack[0];
  if (!top) throw new Error("There is no pls,fix action to undo yet.");

  return Excel.run(async (context) => {
    // Sheet id rather than name, so a rename between action and undo is fine.
    const sheets = top.blocks.map((block) =>
      context.workbook.worksheets.getItemOrNullObject(block.sheetId),
    );
    for (const sheet of sheets) sheet.load("isNullObject");
    await context.sync();

    if (sheets.some((sheet) => sheet.isNullObject)) {
      // Unretryable: this entry can never restore, so it does not stay on
      // the stack the way a merely-failed restore does (see below).
      undoStack = undoStack.slice(1);
      throw new Error("The sheet that action ran on is gone.");
    }

    top.blocks.forEach((block, index) => {
      const range = sheets[index]!.getRange(block.address);
      range.formulas = block.formulas;
      range.numberFormat = block.numberFormat;
      range.setCellProperties(
        block.formats as Excel.SettableCellProperties[][],
      );
    });
    // A sheet protected since the action ran refuses the restore: it comes
    // back named, the way every other write into a locked sheet does.
    await syncWrite(context, "Undo");

    // Only a restore that landed consumes the entry; a failed one stays
    // retryable (this line is unreached when the write above throws).
    undoStack = undoStack.slice(1);
    return undoneMessage(top.label);
  });
}

function undoneMessage(label: string): string {
  const remaining = undoStack.length;
  return remaining > 0
    ? `Undone: ${label}. ${String(remaining)} more to undo.`
    : `Undone: ${label}. Nothing more to undo.`;
}
