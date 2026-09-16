// pls,fix Undo: a five-deep stack, newest first, capturing what a mutating
// action is about to overwrite so it can be restored on demand. Office.js
// writes never reach Excel's own undo stack, so this is the pane's only
// safety net for its last few actions. A capture is pending until the write
// it belongs to lands: syncWrite/paintSync commit it on success and discard
// it on refusal, so a write the host refuses never spends a real slot.

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
// The entry (if any) whose write has not yet settled - set by captureUndoAreas,
// cleared by commitUndo or discardUndo. Never more than one at a time: a new
// capture discards whatever was still pending before pushing its own.
let pendingUndo: UndoEntry | null = null;

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

// Excel for Mac (16.107, proven 16.09) reports an unfilled cell's pattern as
// null and its patternColor as "" rather than the tidy "None" shape a fresh
// clear would carry (lessons.md 2026-08-27: never exact-match state Excel
// gives back), and refuses both verbatim on setCellProperties. Rebuilt as a
// plain "None" fill, or the captured pattern with only the colours it
// actually carries.
function settableFill(
  fill: Excel.CellPropertiesFill | undefined,
): Excel.CellPropertiesFill | undefined {
  if (!fill) return undefined;
  const pattern = fill.pattern as string | null | undefined;
  if (pattern === null || pattern === undefined || pattern === "None") {
    return { pattern: "None" };
  }
  const settable: Excel.CellPropertiesFill = { pattern: fill.pattern };
  if (fill.color) settable.color = fill.color;
  if (fill.patternColor) settable.patternColor = fill.patternColor;
  return settable;
}

const FONT_KEYS = [
  "bold",
  "color",
  "italic",
  "name",
  "size",
  "underline",
] as const;

// Copied key by key rather than spread: an @odata.type annotation on the
// captured object (harmless on Mac, per the diagnosis) never rides along, and
// a captured false or 0 restores instead of being mistaken for "unset".
function settableFont(
  font: Excel.CellPropertiesFont | undefined,
): Excel.CellPropertiesFont | undefined {
  if (!font) return undefined;
  const settable: Excel.CellPropertiesFont = {};
  for (const key of FONT_KEYS) {
    const value = font[key];
    if (value === undefined || value === null) continue;
    (settable as Record<string, unknown>)[key] = value;
  }
  return settable;
}

// Each edge only with the fields Excel actually gave a value for: an empty
// style, weight or colour is the same Mac quirk as an unfilled fill's
// patternColor, so it is dropped rather than sent back verbatim.
function settableBorder(
  border: Excel.CellBorder | undefined,
): Excel.CellBorder | undefined {
  if (!border) return undefined;
  const settable: Excel.CellBorder = {};
  if (border.color) settable.color = border.color;
  if (border.style) settable.style = border.style;
  if (border.weight) settable.weight = border.weight;
  return settable;
}

// Whichever edges the capture carries (top/bottom/left/right/horizontal/
// vertical/diagonalDown/diagonalUp) - requestFormats asks for every edge, but
// the sanitiser makes no assumption about which ones a caller populated.
function settableBorders(
  borders: Excel.CellBorderCollection | undefined,
): Excel.CellBorderCollection | undefined {
  if (!borders) return undefined;
  const settable: Excel.CellBorderCollection = {};
  for (const [edge, border] of Object.entries(borders)) {
    const clean = settableBorder(border as Excel.CellBorder | undefined);
    if (clean) (settable as Record<string, unknown>)[edge] = clean;
  }
  return settable;
}

// The full settable surface a captured cell can carry back through
// setCellProperties, rebuilt field by field so nothing Excel refused on write
// - a null fill pattern, an empty patternColor, an @odata.type annotation -
// ever reaches it again. Font, borders and alignment pass through as
// captured (proven on the Mac): only the empty/null fields are dropped.
export function settableProperties(
  cell: Excel.CellProperties,
): Excel.SettableCellProperties {
  const format = cell.format;
  if (!format) return {};

  const settable: Excel.CellPropertiesFormat = {};
  const fill = settableFill(format.fill);
  if (fill) settable.fill = fill;
  const font = settableFont(format.font);
  if (font) settable.font = font;
  const borders = settableBorders(format.borders);
  if (borders) settable.borders = borders;
  if (format.horizontalAlignment !== undefined) {
    settable.horizontalAlignment = format.horizontalAlignment;
  }
  if (format.verticalAlignment !== undefined) {
    settable.verticalAlignment = format.verticalAlignment;
  }
  if (format.wrapText !== undefined) settable.wrapText = format.wrapText;
  if (format.indentLevel !== undefined) {
    settable.indentLevel = format.indentLevel;
  }
  return { format: settable };
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
  // A capture with no sync after it (its flow threw first) stays pending
  // forever unless the next capture clears it.
  discardUndo();

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

  const captured = ranges.map((range) => {
    const properties = requestFormats(range);
    const sheet = range.worksheet;
    sheet.load("id");
    range.load("formulas,numberFormat");
    return { range, sheet, properties };
  });
  await context.sync();

  const entry: UndoEntry = {
    label: captured.map(({ range }) => range.address).join(", "),
    blocks: captured.map(({ range, sheet, properties }) => ({
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
  pendingUndo = entry;
}

/** The pending capture's write landed: it stays on the stack for good. */
export function commitUndo(): void {
  pendingUndo = null;
}

/**
 * The pending capture's write never landed, so it must not be offered as a
 * restore: dropped from the stack while it is still the top entry (an
 * already-committed entry underneath, or one from an unrelated flow, is never
 * touched) either way.
 */
export function discardUndo(): void {
  if (pendingUndo && undoStack[0] === pendingUndo) {
    undoStack = undoStack.slice(1);
  }
  pendingUndo = null;
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
  // A capture still pending belongs to a flow that never resolved: it must
  // never be offered as a restore.
  discardUndo();
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
        block.formats.map((row) => row.map(settableProperties)),
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
