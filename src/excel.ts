import { type AuditMark, auditGrid } from "./audit";
import { bridgeSeries, cagr, formatCagrLabel } from "./chartmath";
import { type CellClass, classifyCell } from "./classify";
import {
  type BorderSpec,
  buildFillCycle,
  buildFontCycle,
  buildNumberCycles,
  buildRowStyleCycles,
  type CellStyle,
  CLEAR_FILL,
  matchStyleIndex,
  nextInCycle,
  type NumberCycleFamily,
  type RowStyleKind,
  type StyleSpec,
} from "./cycles";
import {
  analyzeGrid,
  type CellValue,
  makeFormatGrid,
  scaleCells,
} from "./model";
import {
  buildCagrFormula,
  detectFillExtent,
  flipSign,
  stepDecimals,
  toggleIfError,
} from "./paste";
import {
  activeTheme,
  currencyNumberFormat,
  getActiveSettings,
  tint,
  type WorkbookTheme,
} from "./settings";

export type PresetName = "title" | "header" | "input" | "formula" | "result";
export type NumberFormatName = "whole" | "decimal" | "currency" | "percent";

export interface SelectionSummary {
  address: string;
  cells: number;
  formulas: number;
  errors: number;
  blanks: number;
}

const staticNumberFormats = {
  whole: "#,##0;[Red](#,##0);-",
  decimal: "#,##0.0;[Red](#,##0.0);-",
  percent: "0.0%;[Red](0.0%);-",
} as const;

function numberFormat(name: NumberFormatName): string {
  if (name === "currency") {
    return currencyNumberFormat(getActiveSettings().currency);
  }
  return staticNumberFormats[name];
}

// ---------------------------------------------------------------------------
// Shared range helpers
// ---------------------------------------------------------------------------

const SELECTION_CELL_CAP = 5_000;
const EDIT_CELL_CAP = 500;
const NO_FILL = "none";
const BASE_WHITE = "#FFFFFF";

// One write per run of same-key cells instead of one per cell: model rows are
// usually uniform, so this keeps the batch small on wide selections. A null key
// leaves the cell untouched.
function writeRuns(
  range: Excel.Range,
  keys: (string | null)[][],
  write: (block: Excel.Range, key: string) => void,
): void {
  keys.forEach((row, rowIndex) => {
    let start = 0;
    while (start < row.length) {
      const key = row[start] ?? null;
      let end = start + 1;
      while (end < row.length && (row[end] ?? null) === key) end += 1;
      if (key !== null) {
        write(
          range.getCell(rowIndex, start).getResizedRange(0, end - start - 1),
          key,
        );
      }
      start = end;
    }
  });
}

// Range addresses arrive sheet-qualified; worksheet.getRange wants the local part.
export function parseAddress(address: string): {
  sheet: string;
  address: string;
} {
  const cut = address.lastIndexOf("!");
  if (cut < 0) return { sheet: "", address };
  return {
    sheet: address.slice(0, cut).replace(/^'|'$/g, "").replace(/''/g, "'"),
    address: address.slice(cut + 1),
  };
}

// Colour, pattern and pattern colour together, so a modeller's own striped fill
// comes back exactly as it was.
function fillKey(fill: Excel.CellPropertiesFill | undefined): string {
  const pattern = fill?.pattern ?? Excel.FillPattern.none;
  if (pattern === Excel.FillPattern.none) return NO_FILL;
  return [
    pattern,
    fill?.color ?? BASE_WHITE,
    fill?.patternColor ?? BASE_WHITE,
  ].join("|");
}

function applyFillKey(block: Excel.Range, key: string): void {
  const { fill } = block.format;
  if (key === NO_FILL) {
    fill.clear();
    return;
  }
  const [pattern, color, patternColor] = key.split("|");
  // Colour first: setting it on an unfilled cell would otherwise force Solid.
  fill.color = color ?? BASE_WHITE;
  fill.pattern = pattern as Excel.FillPattern;
  fill.patternColor = patternColor ?? BASE_WHITE;
}

// ---------------------------------------------------------------------------
// SMT Undo (one slot, restored on demand)
// ---------------------------------------------------------------------------

interface UndoSlot {
  sheetId: string;
  address: string;
  label: string;
  formulas: (string | number | boolean)[][];
  numberFormat: string[][];
  formats: string[][];
}

let undoSlot: UndoSlot | null = null;

// Fill, font colour and weight in one key, so a run of identical cells is a
// single write on the way back.
function formatKey(cell: Excel.CellProperties | undefined): string {
  const font = cell?.format?.font;
  return [fillKey(cell?.format?.fill), font?.color ?? "", font?.bold ? "1" : "0"]
    .join("~");
}

function applyFormatKey(block: Excel.Range, key: string): void {
  const [fill, fontColor, bold] = key.split("~");
  applyFillKey(block, fill ?? NO_FILL);
  if (fontColor) block.format.font.color = fontColor;
  block.format.font.bold = bold === "1";
}

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
  if (range.rowCount * range.columnCount > SELECTION_CELL_CAP) return;

  const properties = range.getCellProperties({
    format: {
      fill: { color: true, pattern: true, patternColor: true },
      font: { color: true, bold: true },
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
    formats: properties.value.map((row) => row.map(formatKey)),
  };
}

export function undoTarget(): string | null {
  return undoSlot?.label ?? null;
}

export async function undoLastAction(): Promise<string> {
  const slot = undoSlot;
  if (!slot) throw new Error("There is no Model Tools action to undo yet.");

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
    writeRuns(range, slot.formats, applyFormatKey);
    await context.sync();

    // Only a restore that landed consumes the slot; a failed one stays retryable.
    undoSlot = null;
    return slot.label;
  });
}

// ---------------------------------------------------------------------------
// Selection and formatting
// ---------------------------------------------------------------------------

export async function inspectSelection(): Promise<SelectionSummary> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address,rowCount,columnCount,formulas,values");
    await context.sync();

    const summary = analyzeGrid(
      range.formulas as CellValue[][],
      range.values as CellValue[][],
    );

    return {
      address: range.address,
      ...summary,
    };
  });
}

export async function applyPreset(name: PresetName): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    await captureUndo(context, range);
    const { format } = range;

    const theme = activeTheme();
    format.font.name = getActiveSettings().font;
    format.font.size = 10;
    format.font.bold = false;
    format.font.italic = false;
    format.font.color = theme.formulaFont;
    format.fill.clear();
    format.horizontalAlignment = Excel.HorizontalAlignment.left;
    format.verticalAlignment = Excel.VerticalAlignment.center;

    switch (name) {
      case "title":
        format.fill.color = theme.titleFill;
        format.font.color = theme.titleText;
        format.font.size = 15;
        format.font.bold = true;
        format.rowHeight = 25;
        break;
      case "header": {
        format.fill.color = theme.headerFill;
        format.font.bold = true;
        const bottom = format.borders.getItem(Excel.BorderIndex.edgeBottom);
        bottom.style = Excel.BorderLineStyle.continuous;
        bottom.color = theme.headerBorder;
        bottom.weight = Excel.BorderWeight.thin;
        break;
      }
      case "input":
        format.font.color = theme.inputFont;
        break;
      case "formula":
        format.font.color = theme.formulaFont;
        break;
      case "result": {
        format.fill.color = theme.resultFill;
        format.font.bold = true;
        const top = format.borders.getItem(Excel.BorderIndex.edgeTop);
        top.style = Excel.BorderLineStyle.double;
        top.color = theme.resultBorder;
        break;
      }
    }

    await context.sync();
  });
}

export async function clearFormats(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    await captureUndo(context, range);
    range.clear(Excel.ClearApplyTo.formats);
    await context.sync();
  });
}

export async function applyNumberFormat(name: NumberFormatName): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount");
    await context.sync();
    await captureUndo(context, range);

    range.numberFormat = makeFormatGrid(
      range.rowCount,
      range.columnCount,
      numberFormat(name),
    );
    await context.sync();
  });
}

// Cycle state lives in the cell: every run reads the active cell and steps once.
function readFill(cell: Excel.Range): string {
  const { fill } = cell.format;
  // Excel reports white for unfilled cells, so the pattern decides.
  return fill.pattern === Excel.FillPattern.none
    ? CLEAR_FILL
    : fill.color.toUpperCase();
}

function readCellStyle(cell: Excel.Range): CellStyle {
  const fill = readFill(cell);
  return {
    fill: fill === CLEAR_FILL ? null : fill,
    fontColor: cell.format.font.color.toUpperCase(),
    bold: cell.format.font.bold,
  };
}

function applyBorder(
  format: Excel.RangeFormat,
  index: Excel.BorderIndex,
  spec: BorderSpec | null | undefined,
): void {
  if (spec === undefined) return;
  const border = format.borders.getItem(index);
  if (spec === null) {
    border.style = Excel.BorderLineStyle.none;
    return;
  }
  if (spec.style === "double") {
    border.style = Excel.BorderLineStyle.double;
  } else {
    border.style = Excel.BorderLineStyle.continuous;
    border.weight = Excel.BorderWeight.thin;
  }
  border.color = spec.color;
}

function applyStyleSpec(format: Excel.RangeFormat, spec: StyleSpec): void {
  if (spec.fill === CLEAR_FILL) format.fill.clear();
  else if (spec.fill !== undefined) format.fill.color = spec.fill;
  if (spec.fontColor !== undefined) format.font.color = spec.fontColor;
  if (spec.bold !== undefined) format.font.bold = spec.bold;
  applyBorder(format, Excel.BorderIndex.edgeTop, spec.topBorder);
  applyBorder(format, Excel.BorderIndex.edgeBottom, spec.bottomBorder);
}

export async function applyNumberCycle(
  family: NumberCycleFamily,
): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const active = range.getCell(0, 0);
    range.load("rowCount,columnCount");
    active.load("numberFormat");
    await context.sync();
    await captureUndo(context, range);

    const current = active.numberFormat[0]?.[0];
    const next = nextInCycle(
      typeof current === "string" ? current : "",
      buildNumberCycles(getActiveSettings())[family],
    );

    range.numberFormat = makeFormatGrid(range.rowCount, range.columnCount, next);
    await context.sync();
  });
}

export async function applyRowStyleCycle(kind: RowStyleKind): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const active = range.getCell(0, 0);
    range.load("rowCount");
    active.load(
      "format/fill/color,format/fill/pattern,format/font/color,format/font/bold",
    );
    await context.sync();

    const variants = buildRowStyleCycles(getActiveSettings())[kind];
    const index = matchStyleIndex(readCellStyle(active), variants);
    const next = variants[(index + 1) % variants.length];
    await captureUndo(context, range);

    if (next) {
      // Edge borders target the whole range, which would leave interior rows
      // bare in a multi-row selection; row styles are per-row by definition.
      if (range.rowCount > 1 && range.rowCount <= 100) {
        for (let row = 0; row < range.rowCount; row += 1) {
          applyStyleSpec(range.getRow(row).format, next);
        }
      } else {
        applyStyleSpec(range.format, next);
      }
    }

    await context.sync();
  });
}

export async function applyFillCycle(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const active = range.getCell(0, 0);
    active.load("format/fill/color,format/fill/pattern");
    await context.sync();

    const next = nextInCycle(
      readFill(active),
      buildFillCycle(getActiveSettings()),
    );
    await captureUndo(context, range);

    if (next === CLEAR_FILL) range.format.fill.clear();
    else range.format.fill.color = next;

    await context.sync();
  });
}

export async function applyFontColorCycle(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const active = range.getCell(0, 0);
    active.load("format/font/color");
    await context.sync();

    const next = nextInCycle(
      active.format.font.color.toUpperCase(),
      buildFontCycle(getActiveSettings()),
    );
    await captureUndo(context, range);

    range.format.font.color = next;
    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Formula edits
// ---------------------------------------------------------------------------

const FILL_SCAN_LIMIT = 1_000;
const SHEET_ROWS = 1_048_576;
const SHEET_COLUMNS = 16_384;

// Macabacus-style fast fill: the data beside the origin decides how far the
// formula travels, so nobody has to select the block first.
export async function fastFillAuto(direction: "right" | "down"): Promise<void> {
  await Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    const sheet = cell.worksheet;
    cell.load("rowIndex,columnIndex,formulas");
    await context.sync();

    const formula = (cell.formulas as CellValue[][])[0]?.[0] ?? null;
    if (typeof formula !== "string" || !formula.startsWith("=")) {
      throw new Error("The active cell must contain a formula.");
    }

    const { rowIndex, columnIndex } = cell;
    const down = direction === "down";
    const span = Math.min(
      FILL_SCAN_LIMIT,
      down ? SHEET_ROWS - rowIndex : SHEET_COLUMNS - columnIndex,
    );
    const lines = (down ? [columnIndex - 1, columnIndex + 1] : [rowIndex - 1, rowIndex + 1])
      .filter((index) => index >= 0 && index < (down ? SHEET_COLUMNS : SHEET_ROWS))
      .map((index) =>
        down
          ? sheet.getRangeByIndexes(rowIndex, index, span, 1)
          : sheet.getRangeByIndexes(index, columnIndex, 1, span),
      );
    for (const line of lines) line.load("values");
    await context.sync();

    const extent = detectFillExtent(
      lines.map((line) => {
        const values = line.values as CellValue[][];
        return down ? values.map((row) => row[0] ?? null) : values[0] ?? [];
      }),
    );
    if (extent === 0) throw new Error("No neighbor data to size the fill.");

    const destination = down
      ? cell.getResizedRange(extent - 1, 0)
      : cell.getResizedRange(0, extent - 1);
    await captureUndo(context, destination);

    destination.copyFrom(cell, Excel.RangeCopyType.formulas);
    await context.sync();
  });
}

export async function toggleIfErrorGuard(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("formulas");
    await context.sync();
    await captureUndo(context, range);

    range.formulas = toggleIfError(
      range.formulas as CellValue[][],
      "0",
    ) as (string | number | boolean)[][];
    await context.sync();
  });
}

export async function scaleSelection(factor: 1000 | 0.001): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("formulas");
    await context.sync();
    await captureUndo(context, range);

    range.formulas = scaleCells(
      range.formulas as CellValue[][],
      factor,
    ) as (string | number | boolean)[][];
    await context.sync();
  });
}

export async function applySignFlip(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("formulas");
    await context.sync();
    await captureUndo(context, range);

    range.formulas = flipSign(
      range.formulas as CellValue[][],
    ) as (string | number | boolean)[][];
    await context.sync();
  });
}

export async function applyDecimalStep(delta: 1 | -1): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("numberFormat");
    await context.sync();
    await captureUndo(context, range);

    range.numberFormat = (range.numberFormat as CellValue[][]).map((row) =>
      row.map((format) => {
        const current = typeof format === "string" ? format : "General";
        // Excel's own Increase Decimal reads General as "0"; stepDecimals, being
        // a pure format transform, leaves an unnumbered format alone.
        const base = current === "General" && delta === 1 ? "0" : current;
        return stepDecimals(base, delta);
      }),
    );
    await context.sync();
  });
}

export async function insertCagr(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount");
    await context.sync();

    const { rowCount, columnCount } = range;
    const periods = (rowCount === 1 ? columnCount : rowCount) - 1;
    if ((rowCount !== 1 && columnCount !== 1) || periods < 1) {
      throw new Error("Select one row or column with at least two periods.");
    }

    const first = range.getCell(0, 0);
    const last = range.getCell(rowCount - 1, columnCount - 1);
    first.load("address");
    last.load("address");
    await context.sync();

    // The result lands just past the series, where a growth row usually sits.
    const destination =
      rowCount === 1 ? last.getOffsetRange(0, 1) : last.getOffsetRange(1, 0);
    await captureUndo(context, destination);

    destination.numberFormat = [[numberFormat("percent")]];
    destination.formulas = [
      [
        buildCagrFormula(
          parseAddress(first.address).address,
          parseAddress(last.address).address,
          periods,
        ),
      ],
    ];
    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Copy and paste
// ---------------------------------------------------------------------------

export type PasteMode = "values" | "formats" | "formulas" | "transpose";

interface CopySource {
  sheetId: string;
  address: string;
  label: string;
}

let copySource: CopySource | null = null;

export function copySourceLabel(): string | null {
  return copySource?.label ?? null;
}

export async function markCopySource(): Promise<string> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const sheet = range.worksheet;
    sheet.load("id");
    range.load("address");
    await context.sync();

    copySource = {
      sheetId: sheet.id,
      address: parseAddress(range.address).address,
      label: range.address,
    };
    return range.address;
  });
}

// Resolved by sheet id, so renaming the source sheet between copy and paste is fine.
async function openCopySource(
  context: Excel.RequestContext,
): Promise<Excel.Range> {
  const source = copySource;
  if (!source) throw new Error("Mark a copy source first.");

  const sheet = context.workbook.worksheets.getItemOrNullObject(source.sheetId);
  sheet.load("isNullObject");
  await context.sync();

  if (sheet.isNullObject) {
    copySource = null;
    throw new Error("The copy source sheet is gone. Mark a new source.");
  }
  return sheet.getRange(source.address);
}

// Read inside the call, never at module scope: office.js defines the enums, and
// the pane also loads in a plain browser where they do not exist yet.
function pasteCopyType(mode: PasteMode): Excel.RangeCopyType {
  switch (mode) {
    case "values":
      return Excel.RangeCopyType.values;
    case "formats":
      return Excel.RangeCopyType.formats;
    case "formulas":
      return Excel.RangeCopyType.formulas;
    case "transpose":
      return Excel.RangeCopyType.all;
  }
}

export async function pasteSpecial(mode: PasteMode): Promise<void> {
  await Excel.run(async (context) => {
    const from = await openCopySource(context);
    const target = context.workbook.getSelectedRange();
    from.load("rowCount,columnCount");
    target.load("rowCount,columnCount");
    await context.sync();

    // Excel grows a smaller destination to the source shape, so undo has to
    // cover the whole footprint, not just what the user selected.
    const transposed = mode === "transpose";
    const rows = transposed ? from.columnCount : from.rowCount;
    const columns = transposed ? from.rowCount : from.columnCount;
    const footprint = target
      .getCell(0, 0)
      .getResizedRange(
        Math.max(rows, target.rowCount) - 1,
        Math.max(columns, target.columnCount) - 1,
      );
    await captureUndo(context, footprint);

    target.copyFrom(from, pasteCopyType(mode), false, transposed);
    await context.sync();
  });
}

// Excel's own paste rewrites relative references; this one keeps the formula
// text byte for byte, which is what a modeller means by "same formula here".
export async function pastePreserveFormulas(): Promise<void> {
  await Excel.run(async (context) => {
    const from = await openCopySource(context);
    const target = context.workbook.getSelectedRange();
    from.load("rowCount,columnCount,formulas");
    await context.sync();

    const destination = target
      .getCell(0, 0)
      .getResizedRange(from.rowCount - 1, from.columnCount - 1);
    await captureUndo(context, destination);

    destination.formulas = from.formulas;
    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Autocolor
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Autocolor on edit (worksheet onChanged, off by default)
// ---------------------------------------------------------------------------

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
      const range = context.workbook.worksheets
        .getItem(event.worksheetId)
        .getRange(event.address);
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
      editHandler = context.workbook.worksheets.onChanged.add(colorChangedRange);
      await context.sync();
    });
    return;
  }

  const handler = editHandler;
  if (!handler) return;
  editHandler = null;
  // Removal has to run on the context the handler was added in.
  await Excel.run(handler.context, async (context) => {
    handler.remove();
    await context.sync();
  });
}

export function setAutocolorOnEdit(enabled: boolean): Promise<void> {
  // Serialized so a fast toggle can never register the handler twice.
  const task = handlerQueue.then(() => applyEditHandler(enabled));
  handlerQueue = task.catch(() => undefined);
  return task;
}

// ---------------------------------------------------------------------------
// Formula audit overlay
// ---------------------------------------------------------------------------

const LONE_FILL = "#E8B4B4";

interface FillSnapshot {
  sheetId: string;
  address: string;
  cells: string[][];
}

// The overlay owns nothing it did not paint: every fill it covers is stored here
// first and written back verbatim. It stays separate from SMT Undo because the
// overlay is a toggle the modeller turns off again, not an edit to the model.
const fillSnapshots = new Map<string, FillSnapshot>();

function overlayKey(mark: AuditMark, patternColor: string): string | null {
  switch (mark) {
    case "horizontal":
      return [Excel.FillPattern.lightHorizontal, BASE_WHITE, patternColor].join("|");
    case "vertical":
      return [Excel.FillPattern.lightVertical, BASE_WHITE, patternColor].join("|");
    case "both":
      return [Excel.FillPattern.crissCross, BASE_WHITE, patternColor].join("|");
    case "lone":
      return [Excel.FillPattern.solid, LONE_FILL, LONE_FILL].join("|");
    case "none":
      return null;
  }
}

export async function snapshotFills(
  context: Excel.RequestContext,
  range: Excel.Range,
): Promise<void> {
  const properties = range.getCellProperties({
    format: { fill: { color: true, pattern: true, patternColor: true } },
  });
  const sheet = range.worksheet;
  sheet.load("id");
  range.load("address");
  await context.sync();

  const address = parseAddress(range.address).address;
  fillSnapshots.set(`${sheet.id}!${address}`, {
    sheetId: sheet.id,
    address,
    cells: properties.value.map((row) =>
      row.map((cell) => fillKey(cell.format?.fill)),
    ),
  });
}

export async function restoreFills(context: Excel.RequestContext): Promise<void> {
  if (fillSnapshots.size === 0) return;

  const pending = [...fillSnapshots.values()].map((snapshot) => ({
    snapshot,
    // Sheet id rather than name, so a rename between paint and restore is fine.
    sheet: context.workbook.worksheets.getItemOrNullObject(snapshot.sheetId),
  }));
  fillSnapshots.clear();

  for (const { sheet } of pending) sheet.load("isNullObject");
  await context.sync();

  for (const { snapshot, sheet } of pending) {
    if (sheet.isNullObject) continue;
    writeRuns(sheet.getRange(snapshot.address), snapshot.cells, applyFillKey);
  }
  await context.sync();
}

export async function toggleAuditOverlay(): Promise<boolean> {
  return Excel.run(async (context) => {
    const selected = context.workbook.getSelectedRange();
    selected.load("rowCount,columnCount");
    await context.sync();

    // One cell says "check this block", not "check this cell".
    let target = selected;
    if (selected.rowCount === 1 && selected.columnCount === 1) {
      target = selected.getSurroundingRegion();
      target.load("rowCount,columnCount");
      await context.sync();
    }
    if (target.rowCount * target.columnCount > SELECTION_CELL_CAP) {
      throw new Error("The audit overlay supports up to 5,000 cells at once.");
    }

    const sheet = target.worksheet;
    sheet.load("id");
    target.load("address,formulasR1C1");
    await context.sync();

    const key = `${sheet.id}!${parseAddress(target.address).address}`;
    const wasOn = fillSnapshots.has(key);
    // Put old fills back before reading new ones, or the next snapshot would
    // capture our own paint over an overlapping range.
    await restoreFills(context);
    if (wasOn) return false;

    await snapshotFills(context, target);

    const patternColor = tint(getActiveSettings().primary, 0.55);
    const marks = auditGrid(target.formulasR1C1 as CellValue[][]);
    writeRuns(
      target,
      marks.map((row) => row.map((mark) => overlayKey(mark, patternColor))),
      applyFillKey,
    );
    await context.sync();
    return true;
  });
}

// ---------------------------------------------------------------------------
// Smart Track (direct precedents and dependents)
// ---------------------------------------------------------------------------

export type TraceDirection = "precedents" | "dependents";

export interface TraceArea {
  sheet: string;
  address: string;
  cellCount: number;
}

export interface TraceResult {
  origin: string;
  areas: TraceArea[];
}

const TRACE_API_SET: Record<TraceDirection, string> = {
  precedents: "1.12",
  dependents: "1.13",
};

function hostSupports(apiSet: string): boolean {
  const requirements = Office.context?.requirements;
  return requirements ? requirements.isSetSupported("ExcelApi", apiSet) : true;
}

export async function traceActiveCell(
  direction: TraceDirection,
): Promise<TraceResult> {
  return Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.load("address");
    await context.sync();

    // The hosted office.js always defines the method, so the host API set decides.
    const method =
      direction === "precedents" ? "getDirectPrecedents" : "getDirectDependents";
    const callable = (cell as unknown as Record<string, unknown>)[method];
    if (
      typeof callable !== "function" ||
      !hostSupports(TRACE_API_SET[direction])
    ) {
      throw new Error("Tracing needs a newer Excel build.");
    }

    const found =
      direction === "precedents"
        ? cell.getDirectPrecedents()
        : cell.getDirectDependents();
    found.ranges.load("items/address,items/cellCount");

    try {
      await context.sync();
    } catch (error) {
      // Excel reports "nothing found" by throwing rather than returning nothing.
      if ((error as { code?: string }).code !== Excel.ErrorCodes.itemNotFound) {
        throw error;
      }
      return { origin: cell.address, areas: [] };
    }

    return {
      origin: cell.address,
      areas: found.ranges.items.map((item) => ({
        ...parseAddress(item.address),
        cellCount: item.cellCount,
      })),
    };
  });
}

export async function selectArea(area: {
  sheet: string;
  address: string;
}): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = area.sheet
      ? context.workbook.worksheets.getItem(area.sheet)
      : context.workbook.worksheets.getActiveWorksheet();
    sheet.activate();
    sheet.getRange(area.address).select();
    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Charts
// ---------------------------------------------------------------------------

const BRIDGE_ROW_CAP = 100;
const CHART_TEXT_SIZE = 9;
const CHART_TITLE_SIZE = 12;
const CAGR_LABEL_WIDTH = 104;
const CAGR_LABEL_HEIGHT = 20;
const CAGR_LABEL_GAP = 8;

// Pies and their relatives have no category or value axis; reading one throws.
const AXIS_FREE_CHARTS = [
  "Pie",
  "PieExploded",
  "PieOfPie",
  "BarOfPie",
  "3DPie",
  "3DPieExploded",
  "Doughnut",
  "DoughnutExploded",
  "Treemap",
  "Sunburst",
  "RegionMap",
];

function chartSeriesColors(): string[] {
  const { primary, accent } = getActiveSettings();
  return [
    accent,
    primary,
    tint(primary, 0.55),
    tint(accent, 0.45),
    tint(primary, 0.78),
    tint(accent, 0.7),
  ];
}

// The brand shell every chart gets: our font everywhere, a bold primary title,
// no gridlines, no chart-area frame, legend under the plot.
function styleChartShell(
  chart: Excel.Chart,
  title: string | null,
  withAxes: boolean,
): void {
  const settings = getActiveSettings();
  const theme = activeTheme();

  chart.format.font.name = settings.font;
  chart.format.font.size = CHART_TEXT_SIZE;
  chart.format.font.color = theme.formulaFont;
  chart.format.border.lineStyle = Excel.ChartLineStyle.none;
  chart.format.roundedCorners = false;

  if (title !== null) chart.title.text = title;
  chart.title.format.font.name = settings.font;
  chart.title.format.font.size = CHART_TITLE_SIZE;
  chart.title.format.font.bold = true;
  chart.title.format.font.color = settings.primary;

  if (withAxes) {
    for (const axis of [chart.axes.categoryAxis, chart.axes.valueAxis]) {
      axis.format.font.name = settings.font;
      axis.format.font.size = CHART_TEXT_SIZE;
      axis.format.font.color = theme.formulaFont;
      axis.majorGridlines.visible = false;
    }
  }

  chart.legend.position = Excel.ChartLegendPosition.bottom;
  chart.legend.overlay = false;
  chart.legend.format.font.name = settings.font;
  chart.legend.format.font.size = CHART_TEXT_SIZE;
  chart.legend.format.font.color = theme.formulaFont;
}

function formatAmount(value: number): string {
  return value.toLocaleString(undefined, { maximumFractionDigits: 1 });
}

// Native Excel waterfall from a two-column bridge table: labels left, values
// right, first and last rows the opening and closing totals.
export async function insertWaterfall(): Promise<string> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const sheet = range.worksheet;
    range.load("rowCount,columnCount,rowIndex,columnIndex,values");
    await context.sync();

    if (range.columnCount !== 2 || range.rowCount < 3) {
      throw new Error(
        "Select two columns, labels and values, with three or more rows.",
      );
    }
    if (range.rowCount > BRIDGE_ROW_CAP) {
      throw new Error("A bridge chart supports up to 100 rows.");
    }

    const values: number[] = [];
    for (const row of range.values as CellValue[][]) {
      const value = row[1];
      if (typeof value !== "number" || !Number.isFinite(value)) {
        throw new Error("The second column must hold numbers only.");
      }
      values.push(value);
    }
    const bridge = bridgeSeries(values);

    // The cell above the table is the chart title when it holds text.
    let heading = "Bridge";
    if (range.rowIndex > 0) {
      const above = sheet.getRangeByIndexes(
        range.rowIndex - 1,
        range.columnIndex,
        1,
        1,
      );
      above.load("values");
      await context.sync();
      const text = (above.values as CellValue[][])[0]?.[0];
      if (typeof text === "string" && text.trim()) heading = text.trim();
    }

    const chart = sheet.charts.add(
      Excel.ChartType.waterfall,
      range,
      Excel.ChartSeriesBy.auto,
    );
    styleChartShell(chart, heading, true);
    chart.legend.visible = false;
    chart.dataLabels.showValue = true;

    const series = chart.series.getItemAt(0);
    series.showConnectorLines = true;
    await context.sync();

    // Office.js has no "set as total" flag for waterfall points, so the opening
    // and closing columns are branded by position instead of by that flag.
    const { primary, accent, external } = getActiveSettings();
    const points = series.points;
    values.forEach((_value, index) => {
      const total = index === 0 || index === values.length - 1;
      const fall = (bridge.fall[index] ?? 0) > 0;
      const color = total ? primary : fall ? external : accent;
      points.getItemAt(index).format.fill.setSolidColor(color);
    });
    await context.sync();

    // Reconciliation: the level the deltas reach against the closing total.
    const last = values.length - 1;
    const implied = (bridge.base[last - 1] ?? 0) + (bridge.rise[last - 1] ?? 0);
    const stated = values[last] ?? 0;
    const ties = Math.abs(implied - stated) <= Math.abs(stated) * 1e-12 + 1e-9;

    if (ties) {
      return `Waterfall added: ${values.length} points, ties at ${formatAmount(stated)}`;
    }
    return `Waterfall added: deltas imply ${formatAmount(implied)}, closing total says ${formatAmount(stated)}`;
  });
}

export async function formatSelectedChart(): Promise<void> {
  await Excel.run(async (context) => {
    const { workbook } = context;
    // The hosted office.js always defines the method, so the host API set decides.
    const callable = (workbook as unknown as Record<string, unknown>)
      .getActiveChartOrNullObject;
    if (typeof callable !== "function" || !hostSupports("1.9")) {
      throw new Error("Chart formatting needs a newer Excel build.");
    }

    const chart = workbook.getActiveChartOrNullObject();
    chart.load("isNullObject");
    await context.sync();
    if (chart.isNullObject) throw new Error("Select a chart first.");

    chart.load("chartType");
    chart.series.load("count");
    await context.sync();

    styleChartShell(chart, null, !AXIS_FREE_CHARTS.includes(String(chart.chartType)));

    const colors = chartSeriesColors();
    for (let index = 0; index < chart.series.count; index += 1) {
      const color = colors[index % colors.length]!;
      chart.series.getItemAt(index).format.fill.setSolidColor(color);
    }
    // One series is already named by the title; its legend is only noise.
    chart.legend.visible = chart.series.count > 1;

    await context.sync();
  });
}

// A floating growth callout beside the series, the way a banker annotates a
// chart by hand. Shapes are worksheet objects, so no range state is touched.
export async function addCagrLabel(): Promise<string> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const sheet = range.worksheet;
    // Range geometry arrived in 1.10; without it the label lands where Excel
    // drops it and the modeller moves it.
    const positioned = hostSupports("1.10");
    range.load("rowCount,columnCount,values");
    if (positioned) range.load("left,top,width");
    await context.sync();

    const { rowCount, columnCount } = range;
    const periods = (rowCount === 1 ? columnCount : rowCount) - 1;
    if ((rowCount !== 1 && columnCount !== 1) || periods < 1) {
      throw new Error("Select one row or column with at least two numbers.");
    }

    const cells = (range.values as CellValue[][]).flat();
    const first = cells[0];
    const last = cells[cells.length - 1];
    if (typeof first !== "number" || typeof last !== "number") {
      throw new Error("The first and last cells must hold numbers.");
    }

    const shapes = (sheet as unknown as { shapes?: Excel.ShapeCollection }).shapes;
    if (!shapes || typeof shapes.addTextBox !== "function" || !hostSupports("1.9")) {
      throw new Error("Chart labels need a newer Excel build.");
    }

    const settings = getActiveSettings();
    const label = formatCagrLabel(cagr(first, last, periods));
    const shape = shapes.addTextBox(label);
    shape.width = CAGR_LABEL_WIDTH;
    shape.height = CAGR_LABEL_HEIGHT;
    if (positioned) {
      shape.left = range.left + range.width + CAGR_LABEL_GAP;
      shape.top = range.top;
    }
    shape.fill.clear();
    shape.lineFormat.visible = false;

    const { textFrame } = shape;
    textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.left;
    textFrame.verticalAlignment = Excel.ShapeTextVerticalAlignment.middle;
    const { font } = textFrame.textRange;
    font.name = settings.font;
    font.size = 11;
    font.bold = true;
    font.color = settings.accent;

    await context.sync();
    return `${label} over ${periods} ${periods === 1 ? "period" : "periods"}`;
  });
}
