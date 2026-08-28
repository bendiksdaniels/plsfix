// Autocolor: classifies cells (input/formula/link/external/partial) and paints
// font color accordingly, plus the color-key legend and the opt-in onChanged
// handler that recolors edited cells live. The handler is off by default.

import { EDIT_CELL_CAP, SELECTION_CELL_CAP, writeRuns } from "./internal";
import { captureUndo } from "./undo";
import { type CellClass, classifyCell } from "../classify";
import { type CellValue } from "../model";
import { activeTheme, getActiveSettings, type WorkbookTheme } from "../settings";

function classFont(kind: CellClass, theme: WorkbookTheme): string | null {
  switch (kind) {
    case "input":
      return theme.inputFont;
    case "formula":
      return theme.formulaFont;
    case "crossSheet":
      return theme.linkFont;
    case "external":
      return theme.externalFont;
    case "partial":
      return theme.partialFont;
    case "blank":
      return null;
  }
}

function colorGrid(
  range: Excel.Range,
  rows: number,
  columns: number,
  formulas: CellValue[][],
  values: CellValue[][],
): void {
  const theme = activeTheme();
  const colors = Array.from({ length: rows }, (_unusedRow, row) =>
    Array.from({ length: columns }, (_unusedColumn, column) =>
      classFont(
        classifyCell(
          formulas[row]?.[column] ?? null,
          values[row]?.[column] ?? null,
        ),
        theme,
      ),
    ),
  );

  writeRuns(range, colors, (block, color) => {
    block.format.font.color = color;
  });
}

export async function autocolorSelection(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount,formulas,values");
    await context.sync();

    if (range.rowCount * range.columnCount > SELECTION_CELL_CAP) {
      throw new Error("Autocolor supports up to 5,000 selected cells at once.");
    }
    await captureUndo(context, range);

    colorGrid(
      range,
      range.rowCount,
      range.columnCount,
      range.formulas as CellValue[][],
      range.values as CellValue[][],
    );

    await context.sync();
  });
}

interface ColorKeyRow {
  label: string;
  example: string | number;
  kind: CellClass;
}

const COLOR_KEY_ROWS: ColorKeyRow[] = [
  { label: "Hardcoded input", example: 1234, kind: "input" },
  { label: "Formula", example: "=A1+B1", kind: "formula" },
  { label: "Cross-sheet link", example: "=Assumptions!B4", kind: "crossSheet" },
  {
    label: "External file link",
    example: "=[Model.xlsx]Sheet1!A1",
    kind: "external",
  },
  { label: "Partial input", example: "=A1*1.05", kind: "partial" },
];

export async function insertColorKey(): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const anchor = context.workbook.getActiveCell();
    anchor.load("rowIndex,columnIndex");
    await context.sync();

    const block = sheet.getRangeByIndexes(
      anchor.rowIndex,
      anchor.columnIndex,
      COLOR_KEY_ROWS.length + 1,
      2,
    );
    await captureUndo(context, block);

    block.format.font.name = getActiveSettings().font;
    block.format.font.size = 10;
    block.format.font.italic = false;
    block.format.verticalAlignment = Excel.VerticalAlignment.center;
    // Text format first, so the example formulas land as text and never compute.
    block.numberFormat = [
      ["@", "@"],
      ...COLOR_KEY_ROWS.map((row) => [
        "@",
        typeof row.example === "number" ? "#,##0" : "@",
      ]),
    ];
    await context.sync();

    const theme = activeTheme();
    const header = block.getRow(0);
    header.values = [["Color key", ""]];
    header.format.fill.color = theme.titleFill;
    header.format.font.color = theme.titleText;
    header.format.font.bold = true;
    header.format.font.size = 11;

    COLOR_KEY_ROWS.forEach((row, index) => {
      const line = block.getRow(index + 1);
      line.values = [[row.label, row.example]];
      line.format.fill.clear();
      line.format.font.bold = false;
      line.format.font.color = classFont(row.kind, theme) ?? theme.formulaFont;
    });

    await context.sync();
  });
}

let editHandler:
  | OfficeExtension.EventHandlerResult<Excel.WorksheetChangedEventArgs>
  | null = null;
let coloringEdit = false;
let handlerQueue: Promise<void> = Promise.resolve();

async function colorChangedRange(
  event: Excel.WorksheetChangedEventArgs,
): Promise<void> {
  // Our own writes must not re-enter, and a removed handler still in flight stops here.
  if (!editHandler || coloringEdit) return;
  coloringEdit = true;

  try {
    await Excel.run(async (context) => {
      // event.address may be sheet-qualified depending on host; the event's own
      // getRange avoids parsing it at all.
      const range = event.getRange(context);
      range.load("rowCount,columnCount");
      await context.sync();

      if (range.rowCount * range.columnCount > EDIT_CELL_CAP) return;

      range.load("formulas,values");
      await context.sync();

      colorGrid(
        range,
        range.rowCount,
        range.columnCount,
        range.formulas as CellValue[][],
        range.values as CellValue[][],
      );
      await context.sync();
    });
  } finally {
    coloringEdit = false;
  }
}

async function applyEditHandler(enabled: boolean): Promise<void> {
  if (enabled) {
    if (editHandler) return;
    await Excel.run(async (context) => {
      // Commit only after the sync that actually registers the handler; a
      // failed sync must not leave a phantom registration behind.
      const handle = context.workbook.worksheets.onChanged.add(colorChangedRange);
      await context.sync();
      editHandler = handle;
    });
    return;
  }

  const handler = editHandler;
  if (!handler) return;
  // Removal has to run on the context the handler was added in. The handle is
  // given up only after the removal has synced — the mirror of the enable
  // path — so a failed sync stays removable instead of orphaning a live
  // registration nothing can reach any more.
  await Excel.run(handler.context, async (context) => {
    handler.remove();
    await context.sync();
  });
  editHandler = null;
}

export function setAutocolorOnEdit(enabled: boolean): Promise<void> {
  // Serialized so a fast toggle can never register the handler twice.
  const task = handlerQueue.then(() => applyEditHandler(enabled));
  handlerQueue = task.catch(() => undefined);
  return task;
}
