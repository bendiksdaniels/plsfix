import { type AuditMark, auditGrid } from "./audit";
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
  wrapFormulasWithIfError,
} from "./model";
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
    context.workbook
      .getSelectedRange()
      .clear(Excel.ClearApplyTo.formats);
    await context.sync();
  });
}

export async function applyNumberFormat(name: NumberFormatName): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount");
    await context.sync();
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

    range.format.font.color = nextInCycle(
      active.format.font.color.toUpperCase(),
      buildFontCycle(getActiveSettings()),
    );
    await context.sync();
  });
}

export async function fastFill(direction: "right" | "down"): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount,formulas");
    await context.sync();

    if (direction === "right" && range.rowCount !== 1) {
      throw new Error("Fill right needs a single-row selection.");
    }
    if (direction === "down" && range.columnCount !== 1) {
      throw new Error("Fill down needs a single-column selection.");
    }

    const source = range.getCell(0, 0);
    source.load("formulas");
    await context.sync();
    const sourceFormula = source.formulas[0]?.[0] as CellValue | undefined;
    if (typeof sourceFormula !== "string" || !sourceFormula.startsWith("=")) {
      throw new Error("The first selected cell must contain a formula.");
    }

    range.copyFrom(source, Excel.RangeCopyType.formulas);
    await context.sync();
  });
}

export async function addIfError(): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("formulas");
    await context.sync();
    range.formulas = wrapFormulasWithIfError(
      range.formulas as CellValue[][],
    ) as (string | number | boolean)[][];
    await context.sync();
  });
}

export async function scaleSelection(factor: 1000 | 0.001): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("formulas");
    await context.sync();
    range.formulas = scaleCells(
      range.formulas as CellValue[][],
      factor,
    ) as (string | number | boolean)[][];
    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Autocolor
// ---------------------------------------------------------------------------

const SELECTION_CELL_CAP = 5_000;
const EDIT_CELL_CAP = 500;

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
const OVERLAY_BASE = "#FFFFFF";
const NO_FILL = "none";

interface FillSnapshot {
  sheetId: string;
  address: string;
  cells: string[][];
}

// The overlay owns nothing it did not paint: every fill it covers is stored here
// first and written back verbatim. C5 generalizes the pair into SMT Undo.
const fillSnapshots = new Map<string, FillSnapshot>();

// Colour, pattern and pattern colour together, so a modeller's own striped fill
// comes back exactly as it was.
function fillKey(fill: Excel.CellPropertiesFill | undefined): string {
  const pattern = fill?.pattern ?? Excel.FillPattern.none;
  if (pattern === Excel.FillPattern.none) return NO_FILL;
  return [
    pattern,
    fill?.color ?? OVERLAY_BASE,
    fill?.patternColor ?? OVERLAY_BASE,
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
  fill.color = color ?? OVERLAY_BASE;
  fill.pattern = pattern as Excel.FillPattern;
  fill.patternColor = patternColor ?? OVERLAY_BASE;
}

function overlayKey(mark: AuditMark, patternColor: string): string | null {
  switch (mark) {
    case "horizontal":
      return [Excel.FillPattern.lightHorizontal, OVERLAY_BASE, patternColor].join("|");
    case "vertical":
      return [Excel.FillPattern.lightVertical, OVERLAY_BASE, patternColor].join("|");
    case "both":
      return [Excel.FillPattern.crissCross, OVERLAY_BASE, patternColor].join("|");
    case "lone":
      return [Excel.FillPattern.solid, LONE_FILL, LONE_FILL].join("|");
    case "none":
      return null;
  }
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
