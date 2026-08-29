// A hand-built Office.js host: an in-memory workbook answering exactly the API
// surface src/excel.ts and src/main.ts touch, so every action can be driven end
// to end without a sideload.
//
// Deliberate simplifications the suite relies on:
//   * context.sync() is a no-op flush: reads are served live from the model, so
//     a stale proxy cannot happen. Load ORDERING is exercised all the same once
//     enableStrictLoadSemantics() is on — see the strict section below.
//   * No formula engine. Writing "=A1+1" stores the text and leaves the cell's
//     value alone; writing a literal sets value and formula together.
//   * formulasR1C1 mirrors the A1 text unless a test seeds it (seedR1C1).
//   * getSelectedRange serves one rectangular area, and throws the way office.js
//     does once helpers.selectAreas() has made the selection multi-area;
//     getSelectedRanges answers areaCount and nothing else.
//   * copyFrom(..., formulas) copies text verbatim; Excel rewrites relative refs.
//   * getImage() hands back a signature-only PNG sized from the fake grid (64pt
//     columns, 20pt rows) or from the requested chart size, never a real picture.
//   * getEntireRow()/getEntireColumn() serve the whole-sheet band Excel does,
//     but a size write spanning more than maxCells rows or columns is refused
//     rather than filling a million map entries.

import { fakePng } from "./fakepng";

export type CellValue = string | number | boolean | null;

export const ROW_LIMIT = 1_048_576;
export const COLUMN_LIMIT = 16_384;
const DEFAULT_ROW_HEIGHT = 15;
const DEFAULT_COLUMN_WIDTH = 64;
const POINTS_PER_ROW = 20;

// ---------------------------------------------------------------------------
// Addresses
// ---------------------------------------------------------------------------

export interface Rect {
  row: number;
  col: number;
  rowCount: number;
  colCount: number;
}

export function columnName(index: number): string {
  let name = "";
  let n = index;
  while (n >= 0) {
    name = String.fromCharCode(65 + (n % 26)) + name;
    n = Math.floor(n / 26) - 1;
  }
  return name;
}

export function columnIndex(name: string): number {
  let n = 0;
  for (const char of name.toUpperCase()) n = n * 26 + (char.charCodeAt(0) - 64);
  return n - 1;
}

function parsePart(part: string): { row: number | null; col: number | null } {
  const match = /^([A-Za-z]+)?(\d+)?$/.exec(part);
  if (!match || (!match[1] && !match[2])) {
    throw new Error(`fake host cannot parse address part "${part}"`);
  }
  return {
    row: match[2] ? Number(match[2]) - 1 : null,
    col: match[1] ? columnIndex(match[1]) : null,
  };
}

export function parseA1(address: string): Rect {
  const [first, second] = address.replace(/\$/g, "").trim().split(":");
  const a = parsePart(first ?? "");
  const b = second ? parsePart(second) : a;

  const rowStart = a.row ?? 0;
  const rowEnd = b.row ?? (a.row === null ? ROW_LIMIT - 1 : rowStart);
  const colStart = a.col ?? 0;
  const colEnd = b.col ?? (a.col === null ? COLUMN_LIMIT - 1 : colStart);

  return {
    row: Math.min(rowStart, rowEnd),
    col: Math.min(colStart, colEnd),
    rowCount: Math.abs(rowEnd - rowStart) + 1,
    colCount: Math.abs(colEnd - colStart) + 1,
  };
}

export function formatA1(rect: Rect): string {
  const fullColumn = rect.rowCount >= ROW_LIMIT;
  const fullRow = rect.colCount >= COLUMN_LIMIT;
  if (fullColumn && !fullRow) {
    const last = columnName(rect.col + rect.colCount - 1);
    return `${columnName(rect.col)}:${last}`;
  }
  if (fullRow && !fullColumn) {
    return `${rect.row + 1}:${rect.row + rect.rowCount}`;
  }
  const start = `${columnName(rect.col)}${rect.row + 1}`;
  if (rect.rowCount === 1 && rect.colCount === 1) return start;
  const end = `${columnName(rect.col + rect.colCount - 1)}${rect.row + rect.rowCount}`;
  return `${start}:${end}`;
}

// The form Excel stores a defined name in: every column letter and row number
// pinned, so inserting rows above the anchor moves it instead of breaking it.
export function absoluteA1(rect: Rect): string {
  return formatA1(rect)
    .split(":")
    .map((part) => part.replace(/([A-Z]+)/, "$$$1").replace(/(\d+)/, "$$$1"))
    .join(":");
}

// Excel quotes a sheet name that is not a bare identifier and doubles apostrophes.
export function quoteSheet(name: string): string {
  return /^[A-Za-z0-9_]+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

// ---------------------------------------------------------------------------
// Cell model
// ---------------------------------------------------------------------------

export interface FakeFont {
  name: string;
  size: number;
  bold: boolean;
  italic: boolean;
  color: string;
  underline: string;
}

export interface FakeFill {
  color: string;
  pattern: string;
  patternColor: string;
}

export interface FakeBorder {
  style: string;
  color: string;
  weight: string;
}

export type BorderEdge =
  "top" | "bottom" | "left" | "right" | "diagonalDown" | "diagonalUp";

const BORDER_EDGES: BorderEdge[] = [
  "top",
  "bottom",
  "left",
  "right",
  "diagonalDown",
  "diagonalUp",
];

export interface FakeHyperlink {
  documentReference?: string;
  address?: string;
  textToDisplay?: string;
  screenTip?: string;
}

// Excel's own default style: the one every untouched cell wears, and the one
// entry a workbook's style table can never be without.
export const NORMAL_STYLE = "Normal";

export interface FakeCell {
  value: CellValue;
  formula: CellValue;
  formulaR1C1: CellValue | null;
  numberFormat: string;
  font: FakeFont;
  fill: FakeFill;
  borders: Record<BorderEdge, FakeBorder>;
  horizontalAlignment: string;
  verticalAlignment: string;
  wrapText: boolean;
  indentLevel: number;
  hyperlink: FakeHyperlink | null;
  // The cell style worn by the cell; every cell wears one, "Normal" by default.
  style: string;
}

export function defaultCell(): FakeCell {
  return {
    value: "",
    formula: "",
    formulaR1C1: null,
    numberFormat: "General",
    font: {
      name: "Calibri",
      size: 11,
      bold: false,
      italic: false,
      color: "#000000",
      underline: "None",
    },
    fill: { color: "#FFFFFF", pattern: "None", patternColor: "#FFFFFF" },
    borders: Object.fromEntries(
      BORDER_EDGES.map((edge) => [
        edge,
        { style: "None", color: "#000000", weight: "Thin" },
      ]),
    ) as Record<BorderEdge, FakeBorder>,
    horizontalAlignment: "General",
    verticalAlignment: "Bottom",
    wrapText: false,
    indentLevel: 0,
    hyperlink: null,
    style: NORMAL_STYLE,
  };
}

const DEFAULT_CELL_JSON = JSON.stringify(defaultCell());

function clone<T>(value: T): T {
  return value === undefined ? value : (JSON.parse(JSON.stringify(value)) as T);
}

function isDefaultCell(cell: FakeCell): boolean {
  return JSON.stringify(cell) === DEFAULT_CELL_JSON;
}

export class FakeSheet {
  cells = new Map<string, FakeCell>();
  rowHeights = new Map<number, number>();
  columnWidths = new Map<number, number>();
  showGridlines = true;
  // Stands in for a recalculation: the last value seen for a formula text, so a
  // formula written back over a clobbered cell shows its result again. Kept off
  // the cell record, which is what tests deep-compare.
  formulaValues = new Map<string, Map<string, CellValue>>();

  constructor(
    public name: string,
    public id: string,
    public visibility = "Visible",
    public position = 0,
  ) {}

  key(row: number, col: number): string {
    return `${row},${col}`;
  }

  // Reads never materialize a cell; only writes do.
  peek(row: number, col: number): FakeCell {
    return this.cells.get(this.key(row, col)) ?? defaultCell();
  }

  edit(row: number, col: number): FakeCell {
    const key = this.key(row, col);
    let cell = this.cells.get(key);
    if (!cell) {
      cell = defaultCell();
      this.cells.set(key, cell);
    }
    return cell;
  }

  remember(
    row: number,
    col: number,
    formula: CellValue,
    value: CellValue,
  ): void {
    if (typeof formula !== "string" || !formula.startsWith("=")) return;
    const key = this.key(row, col);
    let memo = this.formulaValues.get(key);
    if (!memo) {
      memo = new Map();
      this.formulaValues.set(key, memo);
    }
    memo.set(formula, value);
  }

  recall(row: number, col: number, formula: CellValue): CellValue | undefined {
    if (typeof formula !== "string") return undefined;
    return this.formulaValues.get(this.key(row, col))?.get(formula);
  }

  prune(row: number, col: number): void {
    const key = this.key(row, col);
    const cell = this.cells.get(key);
    if (cell && isDefaultCell(cell)) this.cells.delete(key);
  }

  usedRect(): Rect | null {
    let top = Infinity;
    let left = Infinity;
    let bottom = -1;
    let right = -1;
    for (const [key, cell] of this.cells) {
      if (isDefaultCell(cell)) continue;
      const [row, col] = key.split(",").map(Number) as [number, number];
      top = Math.min(top, row);
      left = Math.min(left, col);
      bottom = Math.max(bottom, row);
      right = Math.max(right, col);
    }
    if (bottom < 0) return null;
    return {
      row: top,
      col: left,
      rowCount: bottom - top + 1,
      colCount: right - left + 1,
    };
  }
}

export interface FakeAxis {
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  majorGridlines?: boolean;
  reversePlotOrder?: boolean;
}

export interface FakeSeries {
  showConnectorLines?: boolean;
  fillColor?: string;
  overlap?: number;
  gapWidth?: number;
  pointColors: Record<number, string>;
}

export interface FakeChart {
  sheetName: string;
  name: string;
  width: number;
  height: number;
  chartType: string;
  sourceAddress: string;
  seriesBy: string;
  title: string | null;
  titleFont: Partial<FakeFont>;
  font: Partial<FakeFont>;
  borderLineStyle?: string;
  roundedCorners?: boolean;
  legend: {
    position?: string;
    overlay?: boolean;
    visible?: boolean;
    font: Partial<FakeFont>;
  };
  axes: { category: FakeAxis; value: FakeAxis };
  dataLabels: { showValue?: boolean };
  seriesCount: number;
  series: FakeSeries[];
}

const DEFAULT_CHART_WIDTH = 480;
const DEFAULT_CHART_HEIGHT = 288;

// One place the chart defaults live, so a new field cannot be forgotten at one
// of the three sites that mint a chart record.
export function newChart(
  sheetName: string,
  name: string,
  over: Partial<FakeChart> = {},
): FakeChart {
  return {
    sheetName,
    name,
    width: DEFAULT_CHART_WIDTH,
    height: DEFAULT_CHART_HEIGHT,
    chartType: "",
    sourceAddress: "",
    seriesBy: "",
    title: null,
    titleFont: {},
    font: {},
    legend: { font: {} },
    axes: { category: {}, value: {} },
    dataLabels: {},
    seriesCount: 0,
    series: [],
    ...over,
  };
}

export interface FakeShape {
  sheetName: string;
  text: string;
  width?: number;
  height?: number;
  left?: number;
  top?: number;
  fillCleared: boolean;
  lineVisible?: boolean;
  textFrame: {
    horizontalAlignment?: string;
    verticalAlignment?: string;
    font: Partial<FakeFont>;
  };
}

export interface FakeName {
  name: string;
  formula: string;
  visible: boolean;
}

export interface FakeStyle {
  name: string;
  builtIn: boolean;
}

export interface FakeCommentReply {
  content: string;
  author: string;
}

// A comment belongs to one cell and carries its thread with it; the address is
// sheet-qualified because workbook.comments spans every sheet.
export interface FakeComment {
  address: string;
  content: string;
  author: string;
  replies: FakeCommentReply[];
}

export interface TraceArea {
  address: string;
  cellCount: number;
}

export type TraceConfig = TraceArea[] | "itemNotFound";

export class FakeWorkbook {
  sheets: FakeSheet[] = [];
  settings = new Map<string, string>();
  // What getFilePropertiesAsync reports; empty means an unsaved workbook.
  fileUrl = "";
  names: FakeName[] = [];
  // Excel ships a style table with every workbook; the built-ins in it cannot
  // be deleted, which is what the scrubber has to leave alone.
  styles: FakeStyle[] = [{ name: NORMAL_STYLE, builtIn: true }];
  comments: FakeComment[] = [];
  charts: FakeChart[] = [];
  shapes: FakeShape[] = [];
  selection: { sheetId: string; rect: Rect } = {
    sheetId: "",
    rect: { row: 0, col: 0, rowCount: 1, colCount: 1 },
  };
  // Every area of a ctrl-clicked selection, the first of which is `selection`.
  // Empty means the ordinary single-area case.
  selectionAreas: { sheetId: string; rect: Rect }[] = [];
  activeCell: { sheetId: string; row: number; col: number } | null = null;
  activeSheetId = "";
  activeChart: FakeChart | null = null;
  precedents = new Map<string, TraceConfig>();
  dependents = new Map<string, TraceConfig>();
  private nextId = 1;

  constructor(sheetNames: string[] = ["Sheet1"]) {
    for (const name of sheetNames) this.addSheet(name);
    const first = this.sheets[0];
    if (first) {
      this.activeSheetId = first.id;
      this.selection = {
        sheetId: first.id,
        rect: { row: 0, col: 0, rowCount: 1, colCount: 1 },
      };
    }
  }

  addSheet(name: string): FakeSheet {
    const sheet = new FakeSheet(
      name,
      `sid-${this.nextId}`,
      "Visible",
      this.sheets.length,
    );
    this.nextId += 1;
    this.sheets.push(sheet);
    this.renumber();
    return sheet;
  }

  deleteSheet(name: string): void {
    const index = this.sheets.findIndex((sheet) => sheet.name === name);
    if (index >= 0) this.sheets.splice(index, 1);
    this.renumber();
  }

  find(nameOrId: string): FakeSheet | undefined {
    return this.sheets.find(
      (sheet) => sheet.name === nameOrId || sheet.id === nameOrId,
    );
  }

  ordered(): FakeSheet[] {
    return [...this.sheets].sort((a, b) => a.position - b.position);
  }

  move(sheet: FakeSheet, position: number): void {
    const order = this.ordered().filter((item) => item !== sheet);
    order.splice(Math.max(0, Math.min(position, order.length)), 0, sheet);
    order.forEach((item, index) => {
      item.position = index;
    });
  }

  renumber(): void {
    this.ordered().forEach((sheet, index) => {
      sheet.position = index;
    });
  }

  selectionAddress(): string {
    const sheet = this.find(this.selection.sheetId);
    if (!sheet) return "";
    return `${quoteSheet(sheet.name)}!${formatA1(this.selection.rect)}`;
  }

  areaCount(): number {
    return Math.max(1, this.selectionAreas.length);
  }
}

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

export function hostError(code: string, message: string): Error {
  const error = new Error(message) as Error & { code: string };
  error.code = code;
  return error;
}

// Excel rewrites a bare currency symbol into a locale-tagged code on read-back.
// Excel tags a bare symbol wherever it stands: before the digits or after them.
function rewriteCurrency(format: string): string {
  return format.replace(
    /(\[[^\]]*\])|([€$£¥₹])(?=[ ;)]|$)/g,
    (_match, bracketed: string | undefined, symbol: string | undefined) =>
      bracketed ?? `[$${symbol}-x-fake]`,
  );
}

export interface FakeHostOptions {
  workbook?: FakeWorkbook;
  sheets?: string[];
  rewriteCurrencyFormats?: boolean;
  isSetSupported?: (set: string, version: string) => boolean;
  maxCells?: number;
  strictLoad?: boolean;
  /** Excel for the web: chart font and corners rejected on chartex charts. */
  chartSurfaceUnsupported?: boolean;
  /** Excel's application-level number separators (ExcelApi 1.11). */
  separators?: { decimal: string; thousands: string };
}

// ---------------------------------------------------------------------------
// Strict load semantics
// ---------------------------------------------------------------------------
//
// Real office.js hands out proxies, not values: reading a scalar property that
// was never asked for with load() and committed by a context.sync() throws
// PropertyNotLoaded. The fake serves reads straight from the model, so that
// whole class of bug is invisible unless this layer is switched on.
//
// Every wrapped object belongs to a ROOT — the object load() is called on.
// Navigation properties (range.format, format.fill, chart.series) need no load
// of their own and carry no state; they extend the root's path, so
// range.load("format/fill/color") unlocks range.format.fill.color even though
// the fake builds a fresh child object on every access. What a METHOD hands
// back (getRange, getCell, getItemOrNullObject) is a new root with nothing
// loaded, which is what makes a second Excel.run start from scratch.

const RAW = Symbol("fakehost.raw");

interface Shape {
  // Property reads that need a load plus a sync first.
  scalars?: readonly string[];
  // Navigation property -> kind of the child it hands back.
  children?: Readonly<Record<string, string>>;
  // Method -> kind of the fresh root it hands back.
  returns?: Readonly<Record<string, string>>;
  // Kind of the elements behind a collection's items array.
  items?: string;
  // ClientResult: .value arrives with the next sync, no load needed.
  result?: boolean;
}

// Every kind the add-in can reach, down to the write-only chart and shape
// surfaces: a property named nowhere here would slip through unpoliced, so the
// scalar lists carry the office.js properties even where nothing reads them yet.
const SHAPES: Record<string, Shape> = {
  context: { children: { workbook: "workbook" } },
  workbook: {
    children: {
      worksheets: "worksheets",
      settings: "settings",
      names: "names",
      styles: "styles",
      comments: "comments",
    },
    returns: {
      getSelectedRange: "range",
      getSelectedRanges: "rangeAreas",
      getActiveCell: "range",
      getActiveChartOrNullObject: "chart",
    },
  },
  rangeAreas: { scalars: ["areaCount"] },
  worksheets: {
    scalars: ["items"],
    items: "worksheet",
    children: { onChanged: "events" },
    returns: {
      getItem: "worksheet",
      getItemOrNullObject: "worksheet",
      getActiveWorksheet: "worksheet",
      add: "worksheet",
    },
  },
  // An event source hands back a registration, not a loadable object.
  events: {},
  worksheet: {
    scalars: [
      "id",
      "name",
      "visibility",
      "position",
      "showGridlines",
      "isNullObject",
    ],
    children: { charts: "charts", shapes: "shapes" },
    returns: {
      getRange: "range",
      getRangeByIndexes: "range",
      getUsedRangeOrNullObject: "range",
    },
  },
  charts: { returns: { add: "chart", getItemOrNullObject: "chart" } },
  shapes: { returns: { addTextBox: "shape" } },
  range: {
    scalars: [
      "address",
      "rowIndex",
      "columnIndex",
      "rowCount",
      "columnCount",
      "cellCount",
      "left",
      "top",
      "width",
      "height",
      "values",
      "formulas",
      "formulasR1C1",
      "numberFormat",
      "isNullObject",
    ],
    children: { worksheet: "worksheet", format: "rangeFormat" },
    returns: {
      getCell: "range",
      getRow: "range",
      getColumn: "range",
      getEntireRow: "range",
      getEntireColumn: "range",
      getResizedRange: "range",
      getOffsetRange: "range",
      getSurroundingRegion: "range",
      getUsedRangeOrNullObject: "range",
      getCellProperties: "clientResult",
      getImage: "clientResult",
      getDirectPrecedents: "trace",
      getDirectDependents: "trace",
    },
  },
  rangeFormat: {
    scalars: [
      "horizontalAlignment",
      "verticalAlignment",
      "wrapText",
      "indentLevel",
      "rowHeight",
      "columnWidth",
    ],
    children: { font: "font", fill: "fill", borders: "borders" },
  },
  font: { scalars: ["name", "size", "bold", "italic", "color", "underline"] },
  fill: { scalars: ["color", "pattern", "patternColor"] },
  borders: { returns: { getItem: "border" } },
  border: { scalars: ["style", "color", "weight"] },
  clientResult: { scalars: ["value"], result: true },
  trace: { children: { ranges: "rangeCollection" } },
  rangeCollection: { scalars: ["items"], items: "traceArea" },
  traceArea: { scalars: ["address", "cellCount"] },
  chart: {
    scalars: ["chartType", "name", "width", "height", "isNullObject"],
    returns: { getImage: "clientResult" },
    children: {
      worksheet: "worksheet",
      format: "chartFormat",
      title: "chartTitle",
      axes: "chartAxes",
      legend: "chartLegend",
      dataLabels: "chartDataLabels",
      series: "chartSeries",
    },
  },
  chartFormat: {
    scalars: ["roundedCorners"],
    children: { font: "chartFont", border: "chartBorder", fill: "chartFill" },
  },
  chartFont: {
    scalars: ["name", "size", "bold", "italic", "color", "underline"],
  },
  chartBorder: { scalars: ["lineStyle", "color", "weight"] },
  chartTitle: {
    scalars: ["text", "visible", "overlay"],
    children: { format: "chartFormat" },
  },
  chartAxes: {
    children: { categoryAxis: "chartAxis", valueAxis: "chartAxis" },
  },
  chartAxis: {
    scalars: ["reversePlotOrder"],
    children: { format: "chartFormat", majorGridlines: "chartGridlines" },
  },
  chartGridlines: { scalars: ["visible"] },
  chartLegend: {
    scalars: ["position", "overlay", "visible"],
    children: { format: "chartFormat" },
  },
  chartDataLabels: { scalars: ["showValue"] },
  chartSeries: {
    scalars: ["count"],
    returns: { getItemAt: "chartSeriesItem" },
  },
  chartSeriesItem: {
    scalars: ["name", "showConnectorLines", "overlap", "gapWidth"],
    children: { format: "chartFormat", points: "chartPoints" },
  },
  chartPoints: { scalars: ["count"], returns: { getItemAt: "chartPoint" } },
  chartPoint: { children: { format: "chartFormat" } },
  chartFill: {},
  shape: {
    scalars: ["id", "name", "width", "height", "left", "top", "visible"],
    children: {
      fill: "shapeFill",
      lineFormat: "shapeLine",
      textFrame: "shapeTextFrame",
    },
  },
  shapeFill: { scalars: ["foregroundColor", "transparency"] },
  shapeLine: { scalars: ["visible", "color", "weight"] },
  shapeTextFrame: {
    scalars: ["horizontalAlignment", "verticalAlignment"],
    children: { textRange: "shapeTextRange" },
  },
  shapeTextRange: { scalars: ["text"], children: { font: "chartFont" } },
  settings: { returns: { add: "setting", getItemOrNullObject: "setting" } },
  setting: { scalars: ["key", "value", "isNullObject"] },
  names: {
    scalars: ["items"],
    items: "namedItem",
    returns: {
      add: "namedItem",
      getItem: "namedItem",
      getItemOrNullObject: "namedItem",
    },
  },
  namedItem: {
    scalars: ["name", "formula", "visible", "isNullObject"],
    returns: { getRange: "range", getRangeOrNullObject: "range" },
  },
  styles: { scalars: ["items"], items: "style", returns: { getItem: "style" } },
  style: { scalars: ["name", "builtIn"] },
  comments: { scalars: ["items"], items: "comment" },
  comment: {
    scalars: ["id", "content", "authorName", "authorEmail", "resolved"],
    children: { replies: "commentReplies" },
    returns: { getLocation: "range" },
  },
  commentReplies: { scalars: ["items"], items: "commentReply" },
  commentReply: {
    scalars: ["id", "content", "authorName", "authorEmail"],
    returns: { getLocation: "range" },
  },
};

interface LoadState {
  loaded: Set<string>;
  pending: Set<string>;
  // Prefixes a bare load() covered: every direct scalar below them.
  all: Set<string>;
  pendingAll: Set<string>;
  result: boolean;
  pendingResult: boolean;
}

// The message office.js itself throws, so a failure reads like the real crash.
function notLoaded(property: string): Error {
  const error = new Error(
    `The property '${property}' is not available. Before reading the property's ` +
      `value, call the load method on the containing object and call "context.sync()".`,
  ) as Error & { code: string };
  error.name = "RichApi.Error";
  error.code = "PropertyNotLoaded";
  return error;
}

function notSynced(): Error {
  const error = new Error(
    "The value of the result object has not been loaded yet. Call " +
      '"context.sync()" on the associated request context before reading it.',
  ) as Error & { code: string };
  error.name = "RichApi.Error";
  error.code = "ValueNotLoaded";
  return error;
}

function loadPaths(argument: unknown): string[] | null {
  const split = (text: string): string[] =>
    text
      .split(",")
      .map((part) => part.trim())
      .filter(Boolean);

  if (typeof argument === "string") return split(argument);
  if (Array.isArray(argument)) {
    return argument.flatMap((entry) => split(String(entry)));
  }
  if (argument !== null && typeof argument === "object") {
    return loadPaths((argument as { select?: unknown }).select);
  }
  // load() with no argument: every scalar of that object.
  return null;
}

function rawOf<T>(value: T): T {
  if (value !== null && typeof value === "object") {
    const raw = (value as Record<symbol, unknown>)[RAW];
    if (raw !== undefined) return raw as T;
  }
  return value;
}

class StrictLoads {
  private states = new WeakMap<object, LoadState>();
  private queued = new Set<LoadState>();

  // A sync is what makes a requested property readable.
  commit(): void {
    for (const state of this.queued) {
      for (const path of state.pending) state.loaded.add(path);
      for (const prefix of state.pendingAll) state.all.add(prefix);
      if (state.pendingResult) state.result = true;
      this.reset(state);
    }
    this.queued.clear();
  }

  // A failed sync commits nothing: the whole batch never reached the host.
  drop(): void {
    for (const state of this.queued) this.reset(state);
    this.queued.clear();
  }

  // Wraps an object that owns its own load state: the context, and everything
  // a method hands back.
  root<T>(value: T, kind: string): T {
    if (value === null || typeof value !== "object") return value;
    const state = this.state(value as object);
    if (SHAPES[kind]?.result) {
      state.pendingResult = true;
      this.queued.add(state);
    }
    return this.wrap(value, kind, value as object, "") as T;
  }

  private reset(state: LoadState): void {
    state.pending.clear();
    state.pendingAll.clear();
    state.pendingResult = false;
  }

  private state(root: object): LoadState {
    let found = this.states.get(root);
    if (!found) {
      found = {
        loaded: new Set(),
        pending: new Set(),
        all: new Set(),
        pendingAll: new Set(),
        result: false,
        pendingResult: false,
      };
      this.states.set(root, found);
    }
    return found;
  }

  private record(root: object, prefix: string, argument: unknown): void {
    const state = this.state(root);
    const paths = loadPaths(argument);
    if (paths === null) state.pendingAll.add(prefix);
    else for (const path of paths) state.pending.add(prefix + path);
    this.queued.add(state);
  }

  private require(
    root: object,
    path: string,
    property: string,
    collection: boolean,
  ): void {
    const state = this.states.get(root);
    if (state) {
      if (state.loaded.has(path)) return;
      const cut = path.lastIndexOf("/");
      if (state.all.has(cut < 0 ? "" : path.slice(0, cut + 1))) return;
      // "items/name" makes the items array itself readable, as it does in Excel.
      if (collection) {
        for (const loaded of state.loaded) {
          if (loaded.startsWith(`${path}/`)) return;
        }
      }
    }
    throw notLoaded(property);
  }

  private wrap(
    value: unknown,
    kind: string,
    root: object,
    prefix: string,
  ): unknown {
    if (value === null || typeof value !== "object") return value;
    const shape = SHAPES[kind];
    // A kind nobody described would police nothing, which is the one failure
    // this layer must never have: say so instead of passing everything through.
    if (!shape) throw new Error(`fake host has no strict shape for "${kind}"`);

    return new Proxy(value as Record<string, unknown>, {
      get: (target, property, receiver) => {
        if (property === RAW) return target;
        if (typeof property === "symbol") return Reflect.get(target, property);

        const path = prefix + property;
        const childKind = shape.children?.[property];
        if (childKind !== undefined) {
          return this.wrap(
            Reflect.get(target, property),
            childKind,
            root,
            `${path}/`,
          );
        }

        const raw = Reflect.get(target, property);
        if (typeof raw === "function") {
          return this.method(target, property, shape, root, prefix, receiver);
        }

        if (property === "items" && shape.items !== undefined) {
          this.require(root, path, property, true);
          const items = shape.items;
          return (raw as unknown[]).map((item) =>
            this.wrap(item, items, root, `${path}/`),
          );
        }
        if (property === "value" && shape.result) {
          if (!this.state(root).result) throw notSynced();
          return raw;
        }
        if (shape.scalars?.includes(property)) {
          this.require(root, path, property, false);
          return raw;
        }
        // Not part of the office.js surface: the fake's own fields, and the
        // "then" the runtime probes for whenever a proxy is returned from an
        // async function.
        return raw;
      },
      set(target, property, next) {
        // Writes never need a load, and the setter runs against the model.
        return Reflect.set(target, property, next);
      },
    });
  }

  private method(
    target: Record<string, unknown>,
    property: string,
    shape: Shape,
    root: object,
    prefix: string,
    self: unknown,
  ): (...args: unknown[]) => unknown {
    return (...args: unknown[]): unknown => {
      // Host-side calls take objects, not property reads: hand the fake its own
      // unwrapped proxies so its internals are not policed as add-in reads.
      const plain = args.map((argument) => rawOf(argument));
      if (property === "load") {
        this.record(root, prefix, plain[0]);
        return self;
      }
      const result = (target[property] as (...a: unknown[]) => unknown).apply(
        target,
        plain,
      );
      const returnKind = shape.returns?.[property];
      if (returnKind !== undefined) return this.root(result, returnKind);
      return result;
    };
  }
}

let strictByDefault = false;

// Switches strict load semantics on for every host installed afterwards, which
// is how the integration suite runs the whole add-in against real office.js
// load ordering. installFakeHost({ strictLoad }) overrides it per host.
export function enableStrictLoadSemantics(on = true): void {
  strictByDefault = on;
}

// ---------------------------------------------------------------------------
// Enums (string values copied from @types/office-js)
// ---------------------------------------------------------------------------

const FillPattern = {
  none: "None",
  solid: "Solid",
  gray50: "Gray50",
  gray75: "Gray75",
  gray25: "Gray25",
  horizontal: "Horizontal",
  vertical: "Vertical",
  down: "Down",
  up: "Up",
  checker: "Checker",
  semiGray75: "SemiGray75",
  lightHorizontal: "LightHorizontal",
  lightVertical: "LightVertical",
  lightDown: "LightDown",
  lightUp: "LightUp",
  grid: "Grid",
  crissCross: "CrissCross",
  gray16: "Gray16",
  gray8: "Gray8",
  linearGradient: "LinearGradient",
  rectangularGradient: "RectangularGradient",
} as const;

const BorderIndex = {
  edgeTop: "EdgeTop",
  edgeBottom: "EdgeBottom",
  edgeLeft: "EdgeLeft",
  edgeRight: "EdgeRight",
  insideVertical: "InsideVertical",
  insideHorizontal: "InsideHorizontal",
  diagonalDown: "DiagonalDown",
  diagonalUp: "DiagonalUp",
} as const;

const BorderLineStyle = {
  none: "None",
  continuous: "Continuous",
  dash: "Dash",
  dashDot: "DashDot",
  dashDotDot: "DashDotDot",
  dot: "Dot",
  double: "Double",
  slantDashDot: "SlantDashDot",
} as const;

const BorderWeight = {
  hairline: "Hairline",
  thin: "Thin",
  medium: "Medium",
  thick: "Thick",
} as const;

const HorizontalAlignment = {
  general: "General",
  left: "Left",
  center: "Center",
  right: "Right",
  fill: "Fill",
  justify: "Justify",
  centerAcrossSelection: "CenterAcrossSelection",
  distributed: "Distributed",
} as const;

const VerticalAlignment = {
  top: "Top",
  center: "Center",
  bottom: "Bottom",
  justify: "Justify",
  distributed: "Distributed",
} as const;

const RangeCopyType = {
  all: "All",
  formulas: "Formulas",
  values: "Values",
  formats: "Formats",
  link: "Link",
} as const;

const ChartType = {
  columnClustered: "ColumnClustered",
  barClustered: "BarClustered",
  line: "Line",
  pie: "Pie",
  pieExploded: "PieExploded",
  pieOfPie: "PieOfPie",
  barOfPie: "BarOfPie",
  doughnut: "Doughnut",
  treemap: "Treemap",
  sunburst: "Sunburst",
  regionMap: "RegionMap",
  xyscatter: "XYScatter",
  waterfall: "Waterfall",
} as const;

const ChartSeriesBy = {
  auto: "Auto",
  columns: "Columns",
  rows: "Rows",
} as const;

const SheetVisibility = {
  visible: "Visible",
  hidden: "Hidden",
  veryHidden: "VeryHidden",
} as const;

const ClearApplyTo = {
  all: "All",
  formats: "Formats",
  contents: "Contents",
  hyperlinks: "Hyperlinks",
  removeHyperlinks: "RemoveHyperlinks",
  resetContents: "ResetContents",
} as const;

const ShapeTextHorizontalAlignment = {
  left: "Left",
  center: "Center",
  right: "Right",
  justify: "Justify",
  justifyLow: "JustifyLow",
  distributed: "Distributed",
  thaiDistributed: "ThaiDistributed",
} as const;

const ShapeTextVerticalAlignment = {
  top: "Top",
  middle: "Middle",
  bottom: "Bottom",
  justified: "Justified",
  distributed: "Distributed",
} as const;

const ChartLineStyle = {
  none: "None",
  continuous: "Continuous",
  dash: "Dash",
  dashDot: "DashDot",
  dashDotDot: "DashDotDot",
  dot: "Dot",
  grey25: "Grey25",
  grey50: "Grey50",
  grey75: "Grey75",
  automatic: "Automatic",
  roundDot: "RoundDot",
} as const;

const ChartLegendPosition = {
  invalid: "Invalid",
  top: "Top",
  bottom: "Bottom",
  left: "Left",
  right: "Right",
  corner: "Corner",
  custom: "Custom",
} as const;

const ImageFittingMode = {
  fit: "Fit",
  fitAndCenter: "FitAndCenter",
  fill: "Fill",
} as const;

const RangeUnderlineStyle = {
  none: "None",
  single: "Single",
  double: "Double",
  singleAccountant: "SingleAccountant",
  doubleAccountant: "DoubleAccountant",
} as const;

const ErrorCodes = {
  generalException: "GeneralException",
  invalidArgument: "InvalidArgument",
  itemAlreadyExists: "ItemAlreadyExists",
  itemNotFound: "ItemNotFound",
  invalidOperation: "InvalidOperation",
  invalidSelection: "InvalidSelection",
  unsupportedOperation: "UnsupportedOperation",
} as const;

// BorderIndex -> the cell-level edge it writes. The inside indexes are not in
// here: BordersProxy resolves them to the band of cells that carries the
// interior lines, which is not a single edge of the range.
const BORDER_EDGE: Record<string, BorderEdge> = {
  EdgeTop: "top",
  EdgeBottom: "bottom",
  EdgeLeft: "left",
  EdgeRight: "right",
  DiagonalDown: "diagonalDown",
  DiagonalUp: "diagonalUp",
};

// ---------------------------------------------------------------------------
// Runtime
// ---------------------------------------------------------------------------

interface Registration {
  handler: (args: unknown) => unknown;
}

interface PendingEvent {
  kind: "add" | "remove";
  registration: Registration;
}

class FakeRuntime {
  changeHandlers: Registration[] = [];
  actions = new Map<string, (event?: { completed: () => void }) => void>();
  // Every sync attempt, failed ones included: a round trip is a round trip, and
  // the budget tests count them the way the PowerPoint fake does.
  syncs = 0;
  failSync: Error | null = null;
  // Armed by helpers.failNextImage(): the next getImage queues this error, so
  // the render fails at the sync that was going to commit the anchor with it.
  failImage: Error | null = null;
  supported: (set: string, version: string) => boolean;
  rewriteCurrencyFormats: boolean;
  maxCells: number;
  strict: StrictLoads | null;
  chartSurfaceUnsupported: boolean;
  separators: { decimal: string; thousands: string };

  constructor(
    public workbook: FakeWorkbook,
    options: FakeHostOptions,
  ) {
    this.rewriteCurrencyFormats = options.rewriteCurrencyFormats ?? false;
    this.chartSurfaceUnsupported = options.chartSurfaceUnsupported ?? false;
    this.separators = options.separators ?? { decimal: ".", thousands: "," };
    this.maxCells = options.maxCells ?? 250_000;
    this.supported = options.isSetSupported ?? (() => true);
    const strict = options.strictLoad ?? strictByDefault;
    this.strict = strict ? new StrictLoads() : null;
  }

  format(value: string): string {
    return this.rewriteCurrencyFormats ? rewriteCurrency(value) : value;
  }
}

// getImage is queued like any other call: Excel reports the failure on the sync
// that runs the batch, by which time the writes queued beside it have already
// been applied. Consumed by the one call it was armed for.
function queueImageFailure(runtime: FakeRuntime, ctx: FakeContext): void {
  const failure = runtime.failImage;
  if (!failure) return;
  runtime.failImage = null;
  ctx.queueError(failure);
}

// ---------------------------------------------------------------------------
// Range
// ---------------------------------------------------------------------------

class RangeProxy {
  constructor(
    private runtime: FakeRuntime,
    private ctx: FakeContext,
    public sheet: FakeSheet,
    public rect: Rect,
  ) {}

  load(): this {
    return this;
  }

  get address(): string {
    return `${quoteSheet(this.sheet.name)}!${formatA1(this.rect)}`;
  }

  get rowIndex(): number {
    return this.rect.row;
  }

  get columnIndex(): number {
    return this.rect.col;
  }

  get rowCount(): number {
    return this.rect.rowCount;
  }

  get columnCount(): number {
    return this.rect.colCount;
  }

  get cellCount(): number {
    return this.rect.rowCount * this.rect.colCount;
  }

  get left(): number {
    return this.rect.col * DEFAULT_COLUMN_WIDTH;
  }

  get top(): number {
    return this.rect.row * POINTS_PER_ROW;
  }

  get width(): number {
    return this.rect.colCount * DEFAULT_COLUMN_WIDTH;
  }

  get height(): number {
    return this.rect.rowCount * POINTS_PER_ROW;
  }

  get worksheet(): WorksheetProxy {
    return new WorksheetProxy(this.runtime, this.ctx, this.sheet);
  }

  get format(): RangeFormatProxy {
    return new RangeFormatProxy(this.runtime, this.sheet, this.rect);
  }

  // A guard rather than a hang: an accidental whole-column grid read is a
  // production regression worth failing loudly on.
  private guard(what: string): void {
    if (this.cellCount > this.runtime.maxCells) {
      throw new Error(
        `fake host refused a ${this.cellCount}-cell ${what} (${this.address})`,
      );
    }
  }

  private map<T>(read: (cell: FakeCell) => T): T[][] {
    this.guard("read");
    const { row, col, rowCount, colCount } = this.rect;
    return Array.from({ length: rowCount }, (_unusedRow, r) =>
      Array.from({ length: colCount }, (_unusedColumn, c) =>
        read(this.sheet.peek(row + r, col + c)),
      ),
    );
  }

  private each(write: (cell: FakeCell, r: number, c: number) => void): void {
    this.guard("write");
    const { row, col, rowCount, colCount } = this.rect;
    for (let r = 0; r < rowCount; r += 1) {
      for (let c = 0; c < colCount; c += 1) {
        write(this.sheet.edit(row + r, col + c), r, c);
      }
    }
  }

  get values(): CellValue[][] {
    return this.map((cell) => cell.value);
  }

  set values(grid: CellValue[][]) {
    this.each((cell, r, c) => {
      const entry = grid[r]?.[c];
      if (entry === undefined) return;
      cell.value = entry;
      cell.formula = entry;
      cell.formulaR1C1 = null;
    });
  }

  get formulas(): CellValue[][] {
    return this.map((cell) => cell.formula);
  }

  // No formula engine: a formula write stores the text and the last result seen
  // for it stands in for a recalculation; a literal write sets both, as Excel does.
  set formulas(grid: CellValue[][]) {
    this.each((cell, r, c) => {
      const entry = grid[r]?.[c];
      if (entry === undefined) return;
      cell.formula = entry;
      cell.formulaR1C1 = null;
      if (typeof entry === "string" && entry.startsWith("=")) {
        const known = this.sheet.recall(
          this.rect.row + r,
          this.rect.col + c,
          entry,
        );
        if (known !== undefined) cell.value = known;
        return;
      }
      cell.value = entry;
    });
  }

  get formulasR1C1(): CellValue[][] {
    return this.map((cell) => cell.formulaR1C1 ?? cell.formula);
  }

  set formulasR1C1(grid: CellValue[][]) {
    this.each((cell, r, c) => {
      const entry = grid[r]?.[c];
      if (entry === undefined) return;
      cell.formulaR1C1 = entry;
    });
  }

  get numberFormat(): CellValue[][] {
    return this.map((cell) => cell.numberFormat);
  }

  set numberFormat(grid: CellValue[][]) {
    this.each((cell, r, c) => {
      const entry = grid[r]?.[c];
      if (entry === undefined) return;
      cell.numberFormat = this.runtime.format(String(entry));
    });
  }

  set hyperlink(link: FakeHyperlink) {
    this.each((cell) => {
      cell.hyperlink = clone(link);
      if (link.textToDisplay !== undefined) {
        cell.value = link.textToDisplay;
        cell.formula = link.textToDisplay;
      }
    });
  }

  private at(rect: Rect): RangeProxy {
    return new RangeProxy(this.runtime, this.ctx, this.sheet, rect);
  }

  getCell(row: number, column: number): RangeProxy {
    return this.at({
      row: this.rect.row + row,
      col: this.rect.col + column,
      rowCount: 1,
      colCount: 1,
    });
  }

  getRow(row: number): RangeProxy {
    return this.at({
      row: this.rect.row + row,
      col: this.rect.col,
      rowCount: 1,
      colCount: this.rect.colCount,
    });
  }

  getColumn(column: number): RangeProxy {
    return this.at({
      row: this.rect.row,
      col: this.rect.col + column,
      rowCount: this.rect.rowCount,
      colCount: 1,
    });
  }

  // The bands a size belongs to: a height is a property of whole rows, a width
  // of whole columns. Only the sheet-level size properties make sense on one -
  // a grid read this wide trips the cell guard, which is the point.
  getEntireRow(): RangeProxy {
    return this.at({ ...this.rect, col: 0, colCount: COLUMN_LIMIT });
  }

  getEntireColumn(): RangeProxy {
    return this.at({ ...this.rect, row: 0, rowCount: ROW_LIMIT });
  }

  getResizedRange(deltaRows: number, deltaColumns: number): RangeProxy {
    return this.at({
      ...this.rect,
      rowCount: Math.max(1, this.rect.rowCount + deltaRows),
      colCount: Math.max(1, this.rect.colCount + deltaColumns),
    });
  }

  getOffsetRange(rowOffset: number, columnOffset: number): RangeProxy {
    return this.at({
      ...this.rect,
      row: this.rect.row + rowOffset,
      col: this.rect.col + columnOffset,
    });
  }

  // Excel's current region: grow the rectangle while any bordering line still
  // holds data.
  getSurroundingRegion(): RangeProxy {
    const filled = (row: number, col: number): boolean => {
      const cell = this.sheet.cells.get(this.sheet.key(row, col));
      if (!cell) return false;
      return (
        (cell.value !== "" && cell.value !== null) ||
        (cell.formula !== "" && cell.formula !== null)
      );
    };

    let { row, col, rowCount, colCount } = this.rect;
    if (rowCount === 1 && colCount === 1 && !filled(row, col)) return this;

    let grew = true;
    while (grew) {
      grew = false;
      const line = (
        r0: number,
        c0: number,
        rn: number,
        cn: number,
      ): boolean => {
        for (let r = r0; r < r0 + rn; r += 1) {
          for (let c = c0; c < c0 + cn; c += 1) {
            if (r >= 0 && c >= 0 && filled(r, c)) return true;
          }
        }
        return false;
      };
      if (row > 0 && line(row - 1, col, 1, colCount)) {
        row -= 1;
        rowCount += 1;
        grew = true;
      }
      if (line(row + rowCount, col, 1, colCount)) {
        rowCount += 1;
        grew = true;
      }
      if (col > 0 && line(row, col - 1, rowCount, 1)) {
        col -= 1;
        colCount += 1;
        grew = true;
      }
      if (line(row, col + colCount, rowCount, 1)) {
        colCount += 1;
        grew = true;
      }
    }
    return this.at({ row, col, rowCount, colCount });
  }

  getUsedRangeOrNullObject(): RangeProxy & { isNullObject: boolean } {
    const used = this.sheet.usedRect();
    const proxy = this.at(used ?? this.rect) as RangeProxy & {
      isNullObject: boolean;
    };
    proxy.isNullObject = used === null;
    return proxy;
  }

  // Excel grows a smaller destination to the source shape and tiles a single
  // source cell across a larger one; both are one modulo away from each other.
  copyFrom(
    source: RangeProxy,
    copyType: string = RangeCopyType.all,
    _skipBlanks = false,
    transpose = false,
  ): void {
    // Kept for signature parity with Range.copyFrom; the fake never skips blanks.
    void _skipBlanks;
    const rows = transpose ? source.rect.colCount : source.rect.rowCount;
    const columns = transpose ? source.rect.rowCount : source.rect.colCount;
    const target = this.at({
      ...this.rect,
      rowCount: Math.max(this.rect.rowCount, rows),
      colCount: Math.max(this.rect.colCount, columns),
    });

    // Snapshot up front: a source that overlaps the destination must not read
    // back what this very copy already wrote.
    const values = source.values;
    const formulas = source.formulas;
    const formats = source.map((cell) => clone(cell));
    const numberFormats = source.numberFormat;

    target.each((cell, r, c) => {
      // Transposed, destination (r, c) reads source (c, r); the modulo tiles a
      // one-cell source across the whole destination, which is what a fast fill is.
      const sr = (transpose ? c : r) % source.rect.rowCount;
      const sc = (transpose ? r : c) % source.rect.colCount;
      const from = formats[sr]?.[sc];
      if (!from) return;
      const value = values[sr]?.[sc];
      const formula = formulas[sr]?.[sc];
      const numberFormat = numberFormats[sr]?.[sc];

      if (copyType === RangeCopyType.values) {
        cell.value = value ?? "";
        cell.formula = value ?? "";
        cell.formulaR1C1 = null;
      }
      if (
        copyType === RangeCopyType.formulas ||
        copyType === RangeCopyType.all
      ) {
        cell.value = value ?? "";
        cell.formula = formula ?? "";
        cell.formulaR1C1 = null;
        target.sheet.remember(
          target.rect.row + r,
          target.rect.col + c,
          cell.formula,
          cell.value,
        );
      }
      if (
        copyType === RangeCopyType.formats ||
        copyType === RangeCopyType.all
      ) {
        cell.font = clone(from.font);
        cell.fill = clone(from.fill);
        cell.borders = clone(from.borders);
        cell.horizontalAlignment = from.horizontalAlignment;
        cell.verticalAlignment = from.verticalAlignment;
        cell.wrapText = from.wrapText;
        cell.indentLevel = from.indentLevel;
        cell.numberFormat = String(numberFormat ?? "General");
      }
      if (copyType === RangeCopyType.all)
        cell.hyperlink = clone(from.hyperlink);
    });
  }

  clear(applyTo: string = ClearApplyTo.all): void {
    this.each((cell) => {
      const fresh = defaultCell();
      if (applyTo === ClearApplyTo.all || applyTo === ClearApplyTo.formats) {
        cell.font = fresh.font;
        cell.fill = fresh.fill;
        cell.borders = fresh.borders;
        cell.horizontalAlignment = fresh.horizontalAlignment;
        cell.verticalAlignment = fresh.verticalAlignment;
        cell.wrapText = fresh.wrapText;
        cell.indentLevel = fresh.indentLevel;
        cell.numberFormat = fresh.numberFormat;
      }
      if (
        applyTo === ClearApplyTo.all ||
        applyTo === ClearApplyTo.contents ||
        applyTo === ClearApplyTo.resetContents
      ) {
        cell.value = fresh.value;
        cell.formula = fresh.formula;
        cell.formulaR1C1 = fresh.formulaR1C1;
      }
      if (
        applyTo === ClearApplyTo.all ||
        applyTo === ClearApplyTo.hyperlinks ||
        applyTo === ClearApplyTo.removeHyperlinks
      ) {
        cell.hyperlink = null;
      }
    });
    const { row, col, rowCount, colCount } = this.rect;
    for (let r = 0; r < rowCount; r += 1) {
      for (let c = 0; c < colCount; c += 1) this.sheet.prune(row + r, col + c);
    }
  }

  getCellProperties(options: Record<string, unknown>): {
    value: Record<string, unknown>[][];
  } {
    const wanted = (options?.format ?? {}) as Record<string, unknown>;
    const pick = <T extends object>(
      source: T,
      flags: Record<string, unknown>,
    ): Partial<T> => {
      const out: Record<string, unknown> = {};
      for (const [key, on] of Object.entries(flags)) {
        if (on) out[key] = clone((source as Record<string, unknown>)[key]);
      }
      return out as Partial<T>;
    };

    const value = this.map((cell) => {
      const format: Record<string, unknown> = {};
      if (wanted.fill) {
        format.fill = pick(cell.fill, wanted.fill as Record<string, unknown>);
      }
      if (wanted.font) {
        format.font = pick(cell.font, wanted.font as Record<string, unknown>);
      }
      if (wanted.borders) {
        const flags = wanted.borders as Record<string, unknown>;
        format.borders = Object.fromEntries(
          BORDER_EDGES.map((edge) => [edge, pick(cell.borders[edge], flags)]),
        );
      }
      if (wanted.horizontalAlignment) {
        format.horizontalAlignment = cell.horizontalAlignment;
      }
      if (wanted.verticalAlignment) {
        format.verticalAlignment = cell.verticalAlignment;
      }
      if (wanted.wrapText) format.wrapText = cell.wrapText;
      if (wanted.indentLevel) format.indentLevel = cell.indentLevel;

      const out: Record<string, unknown> = {};
      if (Object.keys(format).length > 0) out.format = format;
      if (options?.hyperlink) out.hyperlink = clone(cell.hyperlink);
      if (options?.style) out.style = cell.style;
      return out;
    });

    return { value };
  }

  // Partial update: a property the caller left out keeps its current value.
  setCellProperties(grid: (Record<string, unknown> | null)[][]): void {
    this.each((cell, r, c) => {
      const props = grid[r]?.[c];
      if (!props) return;
      const format = props.format as Record<string, unknown> | undefined;
      if (format) {
        if (format.fill) Object.assign(cell.fill, clone(format.fill));
        if (format.font) Object.assign(cell.font, clone(format.font));
        if (format.borders) {
          for (const [edge, spec] of Object.entries(
            format.borders as Record<string, object>,
          )) {
            if (spec && cell.borders[edge as BorderEdge]) {
              Object.assign(cell.borders[edge as BorderEdge], clone(spec));
            }
          }
        }
        if (format.horizontalAlignment !== undefined) {
          cell.horizontalAlignment = String(format.horizontalAlignment);
        }
        if (format.verticalAlignment !== undefined) {
          cell.verticalAlignment = String(format.verticalAlignment);
        }
        if (format.wrapText !== undefined)
          cell.wrapText = Boolean(format.wrapText);
        if (format.indentLevel !== undefined) {
          cell.indentLevel = Number(format.indentLevel);
        }
      }
      if (props.hyperlink !== undefined) {
        cell.hyperlink = clone(props.hyperlink as FakeHyperlink | null);
      }
    });
  }

  // Excel renders the range at its on-screen size; the fake grid is a fixed
  // 64pt column by a 20pt row, so the rectangle is the picture.
  getImage(): { value: string } {
    queueImageFailure(this.runtime, this.ctx);
    return { value: fakePng(this.width, this.height) };
  }

  // Excel cannot select on a sheet that is not the active one, so an add-in
  // has to activate the worksheet first. Modelled, because forgetting it is a
  // silent no-op in a fake that does not care and a real bug in Excel.
  select(): void {
    if (this.runtime.workbook.activeSheetId !== this.sheet.id) {
      throw hostError(
        ErrorCodes.invalidOperation,
        `${this.sheet.name} is not the active sheet.`,
      );
    }
    this.runtime.workbook.selection = {
      sheetId: this.sheet.id,
      rect: { ...this.rect },
    };
    this.runtime.workbook.activeCell = null;
    this.runtime.workbook.activeSheetId = this.sheet.id;
  }

  private trace(source: Map<string, TraceConfig>): {
    ranges: { items: TraceArea[]; load: () => void };
  } {
    const config = source.get(this.address) ?? source.get(formatA1(this.rect));
    if (config === "itemNotFound" || config === undefined) {
      this.ctx.queueError(
        hostError(ErrorCodes.itemNotFound, "The requested item was not found."),
      );
      return { ranges: { items: [], load: () => undefined } };
    }
    return { ranges: { items: clone(config), load: () => undefined } };
  }

  getDirectPrecedents() {
    return this.trace(this.runtime.workbook.precedents);
  }

  getDirectDependents() {
    return this.trace(this.runtime.workbook.dependents);
  }
}

// ---------------------------------------------------------------------------
// Range format
// ---------------------------------------------------------------------------

class RangeFormatProxy {
  constructor(
    private runtime: FakeRuntime,
    private sheet: FakeSheet,
    private rect: Rect,
  ) {}

  private each(write: (cell: FakeCell) => void): void {
    const { row, col, rowCount, colCount } = this.rect;
    if (rowCount * colCount > this.runtime.maxCells) {
      throw new Error(
        `fake host refused a ${rowCount * colCount}-cell format write`,
      );
    }
    for (let r = 0; r < rowCount; r += 1) {
      for (let c = 0; c < colCount; c += 1)
        write(this.sheet.edit(row + r, col + c));
    }
  }

  private first(): FakeCell {
    return this.sheet.peek(this.rect.row, this.rect.col);
  }

  get font(): FontProxy {
    return new FontProxy(
      () => this.first().font,
      (apply) => this.each((cell) => apply(cell.font)),
    );
  }

  get fill(): FillProxy {
    return new FillProxy(
      () => this.first().fill,
      (apply) => this.each((cell) => apply(cell.fill)),
    );
  }

  get borders(): BordersProxy {
    return new BordersProxy(this.runtime, this.sheet, this.rect);
  }

  get horizontalAlignment(): string {
    return this.first().horizontalAlignment;
  }

  set horizontalAlignment(value: string) {
    this.each((cell) => {
      cell.horizontalAlignment = value;
    });
  }

  get verticalAlignment(): string {
    return this.first().verticalAlignment;
  }

  set verticalAlignment(value: string) {
    this.each((cell) => {
      cell.verticalAlignment = value;
    });
  }

  get wrapText(): boolean {
    return this.first().wrapText;
  }

  set wrapText(value: boolean) {
    this.each((cell) => {
      cell.wrapText = value;
    });
  }

  get indentLevel(): number {
    return this.first().indentLevel;
  }

  set indentLevel(value: number) {
    this.each((cell) => {
      cell.indentLevel = value;
    });
  }

  // Sizes are sheet state, so they skip the cell guard in each() - but an
  // entire-row or entire-column band spans the whole sheet, and filling a
  // million map entries would hang instead of failing.
  private band(count: number, unit: string): number {
    if (count > this.runtime.maxCells) {
      throw new Error(`fake host refused a ${count}-${unit} size write`);
    }
    return count;
  }

  // Row height and column width are sheet state, not cell state: pls,fix Undo
  // cannot reach them, which the suite asserts explicitly.
  get rowHeight(): number {
    return this.sheet.rowHeights.get(this.rect.row) ?? DEFAULT_ROW_HEIGHT;
  }

  set rowHeight(value: number) {
    const rows = this.band(this.rect.rowCount, "row");
    for (let r = 0; r < rows; r += 1) {
      this.sheet.rowHeights.set(this.rect.row + r, value);
    }
  }

  get columnWidth(): number {
    return this.sheet.columnWidths.get(this.rect.col) ?? DEFAULT_COLUMN_WIDTH;
  }

  set columnWidth(value: number) {
    const columns = this.band(this.rect.colCount, "column");
    for (let c = 0; c < columns; c += 1) {
      this.sheet.columnWidths.set(this.rect.col + c, value);
    }
  }
}

class FontProxy {
  constructor(
    private read: () => FakeFont,
    private write: (apply: (font: FakeFont) => void) => void,
  ) {}

  get name(): string {
    return this.read().name;
  }
  set name(value: string) {
    this.write((font) => {
      font.name = value;
    });
  }

  get size(): number {
    return this.read().size;
  }
  set size(value: number) {
    this.write((font) => {
      font.size = value;
    });
  }

  get bold(): boolean {
    return this.read().bold;
  }
  set bold(value: boolean) {
    this.write((font) => {
      font.bold = value;
    });
  }

  get italic(): boolean {
    return this.read().italic;
  }
  set italic(value: boolean) {
    this.write((font) => {
      font.italic = value;
    });
  }

  get color(): string {
    return this.read().color;
  }
  set color(value: string) {
    this.write((font) => {
      font.color = value;
    });
  }

  get underline(): string {
    return this.read().underline;
  }
  set underline(value: string) {
    this.write((font) => {
      font.underline = value;
    });
  }
}

class FillProxy {
  constructor(
    private read: () => FakeFill,
    private write: (apply: (fill: FakeFill) => void) => void,
  ) {}

  get color(): string {
    return this.read().color;
  }

  // Setting a colour on an unfilled cell forces Solid, exactly as Excel does.
  set color(value: string) {
    this.write((fill) => {
      fill.color = value;
      fill.pattern = FillPattern.solid;
    });
  }

  get pattern(): string {
    return this.read().pattern;
  }
  set pattern(value: string) {
    this.write((fill) => {
      fill.pattern = value;
    });
  }

  get patternColor(): string {
    return this.read().patternColor;
  }
  set patternColor(value: string) {
    this.write((fill) => {
      fill.patternColor = value;
    });
  }

  // Excel reports an unfilled cell as white with pattern None.
  clear(): void {
    this.write((fill) => {
      fill.color = "#FFFFFF";
      fill.pattern = FillPattern.none;
      fill.patternColor = "#FFFFFF";
    });
  }
}

class BordersProxy {
  constructor(
    private runtime: FakeRuntime,
    private sheet: FakeSheet,
    private rect: Rect,
  ) {}

  // An edge border styles the outer line of the RANGE, so only the cells on
  // that edge carry it. Per-row semantics need a per-row loop in the caller.
  // An inside border draws the lines BETWEEN cells: every row but the last
  // carries a horizontal one on its bottom, every column but the last carries a
  // vertical one on its right. A range with nothing inside it has no such line,
  // and the proxy then reads as unset and ignores writes, as Excel does.
  getItem(index: string): BorderProxy {
    const { row, col, rowCount, colCount } = this.rect;
    switch (index) {
      case BorderIndex.edgeTop:
        return this.at({ row, col, rowCount: 1, colCount }, "top");
      case BorderIndex.edgeBottom:
        return this.at(
          { row: row + rowCount - 1, col, rowCount: 1, colCount },
          "bottom",
        );
      case BorderIndex.edgeLeft:
        return this.at({ row, col, rowCount, colCount: 1 }, "left");
      case BorderIndex.edgeRight:
        return this.at(
          { row, col: col + colCount - 1, rowCount, colCount: 1 },
          "right",
        );
      case BorderIndex.insideHorizontal:
        return this.at(
          rowCount > 1 ? { row, col, rowCount: rowCount - 1, colCount } : null,
          "bottom",
        );
      case BorderIndex.insideVertical:
        return this.at(
          colCount > 1 ? { row, col, rowCount, colCount: colCount - 1 } : null,
          "right",
        );
      default:
        return this.at(this.rect, BORDER_EDGE[index] ?? "top");
    }
  }

  private at(band: Rect | null, edge: BorderEdge): BorderProxy {
    return new BorderProxy(this.runtime, this.sheet, band, edge);
  }
}

class BorderProxy {
  constructor(
    private runtime: FakeRuntime,
    private sheet: FakeSheet,
    // Null for a line the range does not have: the inside of a single cell.
    private band: Rect | null,
    private edge: BorderEdge,
  ) {}

  load(): this {
    return this;
  }

  private write(apply: (border: FakeBorder) => void): void {
    if (!this.band) return;
    const { row, col, rowCount, colCount } = this.band;
    if (rowCount * colCount > this.runtime.maxCells) {
      throw new Error("fake host refused an oversized border write");
    }
    for (let r = 0; r < rowCount; r += 1) {
      for (let c = 0; c < colCount; c += 1) {
        apply(this.sheet.edit(row + r, col + c).borders[this.edge]);
      }
    }
  }

  private read(): FakeBorder {
    if (!this.band) {
      return {
        style: BorderLineStyle.none,
        color: "#000000",
        weight: BorderWeight.thin,
      };
    }
    return this.sheet.peek(this.band.row, this.band.col).borders[this.edge];
  }

  get style(): string {
    return this.read().style;
  }
  set style(value: string) {
    this.write((border) => {
      border.style = value;
    });
  }

  get color(): string {
    return this.read().color;
  }
  set color(value: string) {
    this.write((border) => {
      border.color = value;
    });
  }

  get weight(): string {
    return this.read().weight;
  }
  set weight(value: string) {
    this.write((border) => {
      border.weight = value;
    });
  }
}

// ---------------------------------------------------------------------------
// Charts and shapes
// ---------------------------------------------------------------------------

function refusingChartFont(refuse: () => void): ChartFontProxy {
  const font = new ChartFontProxy({});
  for (const key of ["name", "size", "bold", "color"]) {
    Object.defineProperty(font, key, { set: refuse });
  }
  return font;
}

class ChartFontProxy {
  constructor(private target: Partial<FakeFont>) {}
  set name(value: string) {
    this.target.name = value;
  }
  set size(value: number) {
    this.target.size = value;
  }
  set bold(value: boolean) {
    this.target.bold = value;
  }
  set color(value: string) {
    this.target.color = value;
  }
}

class ChartAxisProxy {
  constructor(private axis: FakeAxis) {}
  get format() {
    const axis = this.axis;
    return {
      font: {
        set name(value: string) {
          axis.fontName = value;
        },
        set size(value: number) {
          axis.fontSize = value;
        },
        set color(value: string) {
          axis.fontColor = value;
        },
      },
    };
  }
  get majorGridlines() {
    const axis = this.axis;
    return {
      set visible(value: boolean) {
        axis.majorGridlines = value;
      },
    };
  }

  set reversePlotOrder(value: boolean) {
    this.axis.reversePlotOrder = value;
  }
}

class ChartProxy {
  isNullObject = false;

  constructor(
    private runtime: FakeRuntime,
    private ctx: FakeContext,
    public record: FakeChart,
  ) {}

  load(): this {
    return this;
  }

  get chartType(): string {
    return this.record.chartType;
  }

  get name(): string {
    return this.record.name;
  }

  set name(value: string) {
    this.record.name = value;
  }

  get width(): number {
    return this.record.width;
  }

  set width(value: number) {
    this.record.width = value;
  }

  get height(): number {
    return this.record.height;
  }

  set height(value: number) {
    this.record.height = value;
  }

  get worksheet(): WorksheetProxy {
    const sheet = this.runtime.workbook.find(this.record.sheetName);
    return sheet
      ? new WorksheetProxy(this.runtime, this.ctx, sheet)
      : nullWorksheet(this.runtime, this.ctx);
  }

  // The fitting mode is signature parity only: the fake draws exactly the size
  // it was asked for.
  getImage(
    width: number,
    height: number,
    _fittingMode?: string,
  ): { value: string } {
    void _fittingMode;
    queueImageFailure(this.runtime, this.ctx);
    return { value: fakePng(width, height) };
  }

  activate(): void {
    const workbook = this.runtime.workbook;
    workbook.activeChart = this.record;
    const sheet = workbook.find(this.record.sheetName);
    if (sheet) workbook.activeSheetId = sheet.id;
  }

  // With chartSurfaceUnsupported the font and the corners are refused the way
  // Excel for the web refuses them on chartex charts: queued, and rejecting
  // the sync that carries them with UnsupportedOperation.
  get format() {
    const record = this.record;
    const refuse = this.runtime.chartSurfaceUnsupported
      ? () =>
          this.ctx.queueError(
            hostError(
              ErrorCodes.unsupportedOperation,
              "This operation is not implemented.",
            ),
          )
      : null;
    return {
      font: refuse
        ? refusingChartFont(refuse)
        : new ChartFontProxy(record.font),
      border: {
        set lineStyle(value: string) {
          record.borderLineStyle = value;
        },
      },
      set roundedCorners(value: boolean) {
        if (refuse) refuse();
        else record.roundedCorners = value;
      },
    };
  }

  get title() {
    const record = this.record;
    return {
      set text(value: string) {
        record.title = value;
      },
      format: { font: new ChartFontProxy(record.titleFont) },
    };
  }

  get axes() {
    return {
      categoryAxis: new ChartAxisProxy(this.record.axes.category),
      valueAxis: new ChartAxisProxy(this.record.axes.value),
    };
  }

  get legend() {
    const record = this.record;
    return {
      set position(value: string) {
        record.legend.position = value;
      },
      set overlay(value: boolean) {
        record.legend.overlay = value;
      },
      set visible(value: boolean) {
        record.legend.visible = value;
      },
      format: { font: new ChartFontProxy(record.legend.font) },
    };
  }

  get dataLabels() {
    const record = this.record;
    return {
      set showValue(value: boolean) {
        record.dataLabels.showValue = value;
      },
    };
  }

  get series() {
    const record = this.record;
    return {
      load: () => undefined,
      get count() {
        return record.seriesCount;
      },
      getItemAt(index: number) {
        let entry = record.series[index];
        if (!entry) {
          entry = { pointColors: {} };
          record.series[index] = entry;
        }
        const series = entry;
        return {
          set showConnectorLines(value: boolean) {
            series.showConnectorLines = value;
          },
          set overlap(value: number) {
            series.overlap = value;
          },
          set gapWidth(value: number) {
            series.gapWidth = value;
          },
          format: {
            fill: {
              setSolidColor(color: string) {
                series.fillColor = color;
              },
            },
          },
          points: {
            getItemAt(point: number) {
              return {
                format: {
                  fill: {
                    setSolidColor(color: string) {
                      series.pointColors[point] = color;
                    },
                  },
                },
              };
            },
          },
        };
      },
    };
  }
}

class ShapeProxy {
  constructor(private record: FakeShape) {}

  set width(value: number) {
    this.record.width = value;
  }
  set height(value: number) {
    this.record.height = value;
  }
  set left(value: number) {
    this.record.left = value;
  }
  set top(value: number) {
    this.record.top = value;
  }

  get fill() {
    const record = this.record;
    return {
      clear() {
        record.fillCleared = true;
      },
    };
  }

  get lineFormat() {
    const record = this.record;
    return {
      set visible(value: boolean) {
        record.lineVisible = value;
      },
    };
  }

  get textFrame() {
    const record = this.record;
    return {
      set horizontalAlignment(value: string) {
        record.textFrame.horizontalAlignment = value;
      },
      set verticalAlignment(value: string) {
        record.textFrame.verticalAlignment = value;
      },
      textRange: { font: new ChartFontProxy(record.textFrame.font) },
    };
  }
}

// ---------------------------------------------------------------------------
// Defined names
// ---------------------------------------------------------------------------

class NamedItemProxy {
  isNullObject = false;

  constructor(
    private runtime: FakeRuntime,
    private ctx: FakeContext,
    private record: FakeName,
  ) {}

  load(): this {
    return this;
  }

  get name(): string {
    return this.record.name;
  }

  get formula(): string {
    return this.record.formula;
  }

  get visible(): boolean {
    return this.record.visible;
  }

  set visible(value: boolean) {
    this.record.visible = value;
  }

  delete(): void {
    const list = this.runtime.workbook.names;
    const index = list.indexOf(this.record);
    if (index >= 0) list.splice(index, 1);
  }

  // A name Excel rewrote to #REF! (the rows under it went away), or one
  // pointing at a sheet that is gone, refers to nothing: that is the null
  // object, not a throw.
  private target(): RangeProxy | null {
    const reference = this.record.formula.replace(/^=/, "");
    if (reference.includes("#REF!")) return null;
    try {
      const { sheet, rect } = resolve(this.runtime.workbook, reference);
      return new RangeProxy(this.runtime, this.ctx, sheet, rect);
    } catch {
      return null;
    }
  }

  getRange(): RangeProxy {
    const range = this.target();
    if (!range) {
      throw hostError(
        ErrorCodes.generalException,
        `${this.record.name} does not refer to a range.`,
      );
    }
    return range;
  }

  getRangeOrNullObject(): RangeProxy & { isNullObject: boolean } {
    const found = this.target();
    const range = (found ??
      new RangeProxy(
        this.runtime,
        this.ctx,
        new FakeSheet("__missing__", "__missing__"),
        { row: 0, col: 0, rowCount: 1, colCount: 1 },
      )) as RangeProxy & { isNullObject: boolean };
    range.isNullObject = found === null;
    return range;
  }
}

// ---------------------------------------------------------------------------
// Comments
// ---------------------------------------------------------------------------

// Where a comment hangs: office.js hands back the one cell it is anchored to,
// never the range that was selected when it was written.
function commentCell(
  runtime: FakeRuntime,
  ctx: FakeContext,
  address: string,
): RangeProxy {
  const { sheet, rect } = resolve(runtime.workbook, address);
  return new RangeProxy(runtime, ctx, sheet, {
    row: rect.row,
    col: rect.col,
    rowCount: 1,
    colCount: 1,
  });
}

class CommentReplyProxy {
  constructor(
    private runtime: FakeRuntime,
    private ctx: FakeContext,
    private record: FakeCommentReply,
    private address: string,
  ) {}

  load(): this {
    return this;
  }

  get content(): string {
    return this.record.content;
  }

  get authorName(): string {
    return this.record.author;
  }

  getLocation(): RangeProxy {
    return commentCell(this.runtime, this.ctx, this.address);
  }
}

class CommentProxy {
  constructor(
    private runtime: FakeRuntime,
    private ctx: FakeContext,
    private record: FakeComment,
  ) {}

  load(): this {
    return this;
  }

  get content(): string {
    return this.record.content;
  }

  get authorName(): string {
    return this.record.author;
  }

  // The thread is a collection of its own: office.js offers no load option for
  // it on the comment collection, so an add-in loads it one level down, through
  // the comment that owns it. Every reply sits on that comment's cell.
  get replies(): { load: () => void; items: CommentReplyProxy[] } {
    const { runtime, ctx, record } = this;
    return {
      load: () => undefined,
      get items(): CommentReplyProxy[] {
        return record.replies.map(
          (reply) => new CommentReplyProxy(runtime, ctx, reply, record.address),
        );
      },
    };
  }

  getLocation(): RangeProxy {
    return commentCell(this.runtime, this.ctx, this.record.address);
  }
}

// ---------------------------------------------------------------------------
// Style
// ---------------------------------------------------------------------------

class StyleProxy {
  constructor(
    private runtime: FakeRuntime,
    private record: FakeStyle,
  ) {}

  load(): this {
    return this;
  }

  get name(): string {
    return this.record.name;
  }

  get builtIn(): boolean {
    return this.record.builtIn;
  }

  // Excel refuses to delete one of its own styles; a scrubber that offered one
  // would fail at the host rather than quietly do nothing.
  delete(): void {
    if (this.record.builtIn) {
      throw hostError(
        ErrorCodes.invalidArgument,
        `${this.record.name} is a built-in style.`,
      );
    }
    const list = this.runtime.workbook.styles;
    const index = list.indexOf(this.record);
    if (index >= 0) list.splice(index, 1);
  }
}

// ---------------------------------------------------------------------------
// Worksheet
// ---------------------------------------------------------------------------

class WorksheetProxy {
  isNullObject = false;

  constructor(
    private runtime: FakeRuntime,
    private ctx: FakeContext,
    public sheet: FakeSheet,
  ) {}

  load(): this {
    return this;
  }

  get id(): string {
    return this.sheet.id;
  }

  get name(): string {
    return this.sheet.name;
  }

  set name(value: string) {
    this.sheet.name = value;
  }

  get visibility(): string {
    return this.sheet.visibility;
  }

  set visibility(value: string) {
    this.sheet.visibility = value;
  }

  get position(): number {
    return this.sheet.position;
  }

  set position(value: number) {
    this.runtime.workbook.move(this.sheet, value);
  }

  get showGridlines(): boolean {
    return this.sheet.showGridlines;
  }

  set showGridlines(value: boolean) {
    this.sheet.showGridlines = value;
  }

  activate(): void {
    this.runtime.workbook.activeSheetId = this.sheet.id;
  }

  getRange(address: string): RangeProxy {
    return new RangeProxy(this.runtime, this.ctx, this.sheet, parseA1(address));
  }

  getRangeByIndexes(
    row: number,
    col: number,
    rowCount: number,
    colCount: number,
  ): RangeProxy {
    return new RangeProxy(this.runtime, this.ctx, this.sheet, {
      row,
      col,
      rowCount,
      colCount,
    });
  }

  getUsedRangeOrNullObject(): RangeProxy & { isNullObject: boolean } {
    const used = this.sheet.usedRect();
    const proxy = new RangeProxy(
      this.runtime,
      this.ctx,
      this.sheet,
      used ?? { row: 0, col: 0, rowCount: 1, colCount: 1 },
    ) as RangeProxy & { isNullObject: boolean };
    proxy.isNullObject = used === null;
    return proxy;
  }

  get charts() {
    const runtime = this.runtime;
    const ctx = this.ctx;
    const sheet = this.sheet;
    return {
      add(chartType: string, source: RangeProxy, seriesBy: string): ChartProxy {
        const record = newChart(
          sheet.name,
          `Chart ${runtime.workbook.charts.length + 1}`,
          {
            chartType,
            sourceAddress: source.address,
            seriesBy,
            seriesCount: Math.max(1, source.columnCount - 1),
          },
        );
        runtime.workbook.charts.push(record);
        runtime.workbook.activeChart = record;
        return new ChartProxy(runtime, ctx, record);
      },
      // Charts are named workbook-wide but looked up sheet by sheet, which is
      // how a chart that was moved is found again.
      getItemOrNullObject(name: string): ChartProxy {
        const record = runtime.workbook.charts.find(
          (chart) => chart.sheetName === sheet.name && chart.name === name,
        );
        if (record) return new ChartProxy(runtime, ctx, record);
        const missing = new ChartProxy(runtime, ctx, newChart("", ""));
        missing.isNullObject = true;
        return missing;
      },
    };
  }

  get shapes() {
    const runtime = this.runtime;
    const sheet = this.sheet;
    return {
      addTextBox(text: string): ShapeProxy {
        const record: FakeShape = {
          sheetName: sheet.name,
          text,
          fillCleared: false,
          textFrame: { font: {} },
        };
        runtime.workbook.shapes.push(record);
        return new ShapeProxy(record);
      },
    };
  }
}

// A worksheet that is not there: same shape, isNullObject true, never touched
// past the check in src/excel.ts.
function nullWorksheet(
  runtime: FakeRuntime,
  ctx: FakeContext,
): WorksheetProxy & { isNullObject: true } {
  const proxy = new WorksheetProxy(
    runtime,
    ctx,
    new FakeSheet("__missing__", "__missing__"),
  ) as WorksheetProxy & { isNullObject: true };
  proxy.isNullObject = true;
  return proxy;
}

// ---------------------------------------------------------------------------
// Workbook and request context
// ---------------------------------------------------------------------------

class WorksheetCollectionProxy {
  constructor(
    private runtime: FakeRuntime,
    private ctx: FakeContext,
  ) {}

  load(): this {
    return this;
  }

  get items(): WorksheetProxy[] {
    return this.runtime.workbook
      .ordered()
      .map((sheet) => new WorksheetProxy(this.runtime, this.ctx, sheet));
  }

  getItem(nameOrId: string): WorksheetProxy {
    const sheet = this.runtime.workbook.find(nameOrId);
    if (!sheet) {
      throw hostError(ErrorCodes.itemNotFound, `No sheet named ${nameOrId}.`);
    }
    return new WorksheetProxy(this.runtime, this.ctx, sheet);
  }

  getItemOrNullObject(nameOrId: string): WorksheetProxy {
    const sheet = this.runtime.workbook.find(nameOrId);
    return sheet
      ? new WorksheetProxy(this.runtime, this.ctx, sheet)
      : nullWorksheet(this.runtime, this.ctx);
  }

  getActiveWorksheet(): WorksheetProxy {
    const workbook = this.runtime.workbook;
    const sheet = workbook.find(workbook.activeSheetId) ?? workbook.sheets[0];
    if (!sheet) throw hostError(ErrorCodes.itemNotFound, "No sheets.");
    return new WorksheetProxy(this.runtime, this.ctx, sheet);
  }

  add(name?: string): WorksheetProxy {
    const workbook = this.runtime.workbook;
    const chosen = name ?? `Sheet${workbook.sheets.length + 1}`;
    if (workbook.find(chosen)) {
      throw hostError(
        ErrorCodes.itemAlreadyExists,
        `${chosen} already exists.`,
      );
    }
    return new WorksheetProxy(
      this.runtime,
      this.ctx,
      workbook.addSheet(chosen),
    );
  }

  // Registration is committed by the sync that follows, so a failed sync leaves
  // no handler behind — the same contract the real host offers.
  get onChanged() {
    const ctx = this.ctx;
    return {
      add(handler: (args: unknown) => unknown) {
        const registration: Registration = { handler };
        ctx.queueEvent({ kind: "add", registration });
        return {
          context: ctx,
          remove() {
            ctx.queueEvent({ kind: "remove", registration });
          },
        };
      },
    };
  }
}

class WorkbookProxy {
  constructor(
    private runtime: FakeRuntime,
    private ctx: FakeContext,
  ) {}

  get worksheets(): WorksheetCollectionProxy {
    return new WorksheetCollectionProxy(this.runtime, this.ctx);
  }

  private sheetOf(id: string): FakeSheet {
    const sheet =
      this.runtime.workbook.find(id) ?? this.runtime.workbook.sheets[0];
    if (!sheet) throw hostError(ErrorCodes.itemNotFound, "No sheets.");
    return sheet;
  }

  // office.js: "If there are multiple ranges selected, this method will throw
  // an error." Queued, so the failure arrives on the next sync with no stage
  // of its own - which is what the flows using getSelectedRanges avoid.
  getSelectedRange(): RangeProxy {
    const workbook = this.runtime.workbook;
    if (workbook.areaCount() > 1) {
      this.ctx.queueError(
        hostError(
          ErrorCodes.invalidSelection,
          "The selected range contains multiple areas.",
        ),
      );
    }
    const { sheetId, rect } = workbook.selection;
    return new RangeProxy(this.runtime, this.ctx, this.sheetOf(sheetId), {
      ...rect,
    });
  }

  getSelectedRanges(): { areaCount: number; load: () => void } {
    return {
      areaCount: this.runtime.workbook.areaCount(),
      load: () => undefined,
    };
  }

  getActiveCell(): RangeProxy {
    const workbook = this.runtime.workbook;
    const active = workbook.activeCell;
    if (active) {
      return new RangeProxy(
        this.runtime,
        this.ctx,
        this.sheetOf(active.sheetId),
        {
          row: active.row,
          col: active.col,
          rowCount: 1,
          colCount: 1,
        },
      );
    }
    const { sheetId, rect } = workbook.selection;
    return new RangeProxy(this.runtime, this.ctx, this.sheetOf(sheetId), {
      row: rect.row,
      col: rect.col,
      rowCount: 1,
      colCount: 1,
    });
  }

  getActiveChartOrNullObject(): ChartProxy {
    const record = this.runtime.workbook.activeChart;
    if (record) return new ChartProxy(this.runtime, this.ctx, record);
    const empty = new ChartProxy(this.runtime, this.ctx, newChart("", ""));
    empty.isNullObject = true;
    return empty;
  }

  get settings() {
    const store = this.runtime.workbook.settings;
    return {
      add(key: string, value: unknown) {
        store.set(
          key,
          typeof value === "string" ? value : JSON.stringify(value),
        );
        return { key, value, load: () => undefined };
      },
      getItemOrNullObject(key: string) {
        const has = store.has(key);
        return {
          isNullObject: !has,
          value: store.get(key) ?? "",
          load: () => undefined,
        };
      },
    };
  }

  get names() {
    const runtime = this.runtime;
    const ctx = this.ctx;
    const list = runtime.workbook.names;
    const wrap = (record: FakeName) => new NamedItemProxy(runtime, ctx, record);
    return {
      load: () => undefined,
      get items() {
        return list.map(wrap);
      },
      // Excel stores a range reference absolutely, so the name survives rows
      // being inserted above it; a string reference is taken as given.
      add(name: string, reference: RangeProxy | string): NamedItemProxy {
        if (list.some((entry) => entry.name === name)) {
          throw hostError(
            ErrorCodes.itemAlreadyExists,
            `${name} already exists.`,
          );
        }
        const formula =
          typeof reference === "string"
            ? reference.startsWith("=")
              ? reference
              : `=${reference}`
            : `=${quoteSheet(reference.sheet.name)}!${absoluteA1(reference.rect)}`;
        const record: FakeName = { name, formula, visible: true };
        list.push(record);
        return wrap(record);
      },
      getItem(name: string): NamedItemProxy {
        const record = list.find((entry) => entry.name === name);
        if (!record) {
          throw hostError(ErrorCodes.itemNotFound, `No name ${name}.`);
        }
        return wrap(record);
      },
      getItemOrNullObject(name: string): NamedItemProxy {
        const record = list.find((entry) => entry.name === name);
        if (record) return wrap(record);
        const missing = wrap({ name, formula: "", visible: false });
        missing.isNullObject = true;
        return missing;
      },
    };
  }

  // Every comment in the workbook, sheet by sheet, the way office.js serves
  // them: the cell each one hangs on comes from getLocation(), not from here.
  get comments() {
    const runtime = this.runtime;
    const ctx = this.ctx;
    const list = runtime.workbook.comments;
    return {
      load: () => undefined,
      get items(): CommentProxy[] {
        return list.map((record) => new CommentProxy(runtime, ctx, record));
      },
    };
  }

  get styles() {
    const runtime = this.runtime;
    const list = runtime.workbook.styles;
    const wrap = (record: FakeStyle) => new StyleProxy(runtime, record);
    return {
      load: () => undefined,
      get items() {
        return list.map(wrap);
      },
      getItem(name: string): StyleProxy {
        const record = list.find((entry) => entry.name === name);
        if (!record) {
          throw hostError(ErrorCodes.itemNotFound, `No style ${name}.`);
        }
        return wrap(record);
      },
    };
  }
}

// Excel.Application: the separators the host is set to, read-only here as in
// Office.js (ExcelApi 1.11).
class ApplicationProxy {
  constructor(private runtime: FakeRuntime) {}

  load(): this {
    return this;
  }

  get decimalSeparator(): string {
    return this.runtime.separators.decimal;
  }

  get thousandsSeparator(): string {
    return this.runtime.separators.thousands;
  }
}

class FakeContext {
  workbook: WorkbookProxy;
  application: ApplicationProxy;
  private pending: PendingEvent[] = [];
  private error: Error | null = null;

  constructor(private runtime: FakeRuntime) {
    this.workbook = new WorkbookProxy(runtime, this);
    this.application = new ApplicationProxy(runtime);
  }

  queueEvent(event: PendingEvent): void {
    this.pending.push(event);
  }

  queueError(error: Error): void {
    this.error = error;
  }

  // No-op flush for values: the model is already current. Pending event
  // registrations, a queued host error and — under strict load semantics — the
  // properties this batch asked for are what a sync really decides.
  async sync(): Promise<void> {
    this.runtime.syncs += 1;
    const queued = this.error ?? this.runtime.failSync;
    if (queued) {
      this.error = null;
      this.runtime.failSync = null;
      this.pending.length = 0;
      this.runtime.strict?.drop();
      throw queued;
    }
    for (const event of this.pending) {
      const handlers = this.runtime.changeHandlers;
      if (event.kind === "add") handlers.push(event.registration);
      else {
        const index = handlers.indexOf(event.registration);
        if (index >= 0) handlers.splice(index, 1);
      }
    }
    this.pending.length = 0;
    this.runtime.strict?.commit();
  }
}

// ---------------------------------------------------------------------------
// Install
// ---------------------------------------------------------------------------

export type SeedEntry =
  CellValue | { value?: CellValue; formula?: CellValue; r1c1?: CellValue };

export interface FakeHelpers {
  sheet(name: string): FakeSheet;
  addSheet(name: string): FakeSheet;
  deleteSheet(name: string): void;
  addName(name: string, formula: string): void;
  // A comment on one cell, plus its thread; a reply keeps the comment's author
  // unless it names its own.
  addComment(
    address: string,
    content: string,
    author: string,
    replies?: { content: string; author?: string }[],
  ): void;
  setNameFormula(name: string, formula: string): void;
  breakName(name: string): void;
  select(address: string): void;
  // A ctrl-clicked selection: the first address is what getSelectedRange would
  // have served, and getSelectedRanges reports one area per address.
  selectAreas(addresses: string[]): void;
  setActiveCell(address: string): void;
  seed(address: string, grid: SeedEntry[][]): void;
  setNumberFormat(address: string, format: string): void;
  setFill(address: string, fill: Partial<FakeFill>): void;
  setFont(address: string, font: Partial<FakeFont>): void;
  addStyle(name: string, builtIn?: boolean): void;
  setStyle(address: string, name: string): void;
  cell(address: string): FakeCell;
  fill(address: string): FakeFill;
  font(address: string): FakeFont;
  border(address: string, edge: BorderEdge): FakeBorder;
  numberFormat(address: string): string;
  value(address: string): CellValue;
  formula(address: string): CellValue;
  cellMap(sheetName: string): Record<string, FakeCell>;
  rowHeight(sheetName: string, row: number): number;
  columnWidth(sheetName: string, col: number): number;
  setPrecedents(address: string, config: TraceConfig): void;
  setDependents(address: string, config: TraceConfig): void;
  addChart(sheetName: string, chart?: Partial<FakeChart>): FakeChart;
  moveChart(name: string, toSheet: string): void;
  setActiveChart(chart: FakeChart | null): void;
  setting(key: string): string | null;
  setSetting(key: string, value: string): void;
  setSupported(check: (set: string, version: string) => boolean): void;
  failNextSync(error?: Error): void;
  // Makes the next Range/Chart getImage fail, the way a chart mid-render or a
  // protected sheet does, without touching the writes queued beside it.
  failNextImage(error?: Error): void;
  changeHandlerCount(): number;
  // Round trips so far, for the tests that hold a flow to a sync budget.
  syncCount(): number;
  // One area ("C5") or several ("C5,H9"), the way a ctrl-clicked edit arrives:
  // every area is sheet-qualified in the address the event carries, exactly as
  // the host writes it.
  fireChanged(sheetIdOrName: string, address: string): Promise<void>;
  actions(): Map<string, (event?: { completed: () => void }) => void>;
}

function resolve(
  workbook: FakeWorkbook,
  address: string,
): { sheet: FakeSheet; rect: Rect } {
  const cut = address.lastIndexOf("!");
  const name =
    cut < 0
      ? ""
      : address.slice(0, cut).replace(/^'|'$/g, "").replace(/''/g, "'");
  const local = cut < 0 ? address : address.slice(cut + 1);
  const sheet = name
    ? workbook.find(name)
    : workbook.find(workbook.activeSheetId);
  if (!sheet) throw new Error(`fake host: no sheet for "${address}"`);
  return { sheet, rect: parseA1(local) };
}

export function installFakeHost(options: FakeHostOptions = {}): {
  workbook: FakeWorkbook;
  helpers: FakeHelpers;
} {
  const workbook = options.workbook ?? new FakeWorkbook(options.sheets);
  const runtime = new FakeRuntime(workbook, options);

  const excel = {
    run(first: unknown, second?: unknown): Promise<unknown> {
      const callback = (typeof first === "function" ? first : second) as (
        context: FakeContext,
      ) => unknown;
      const given = rawOf(first);
      const context =
        given instanceof FakeContext ? given : new FakeContext(runtime);
      const handed = runtime.strict
        ? runtime.strict.root(context, "context")
        : context;
      return Promise.resolve().then(() => callback(handed));
    },
    RequestContext: FakeContext,
    FillPattern,
    BorderIndex,
    BorderLineStyle,
    BorderWeight,
    HorizontalAlignment,
    VerticalAlignment,
    RangeCopyType,
    ChartType,
    ChartSeriesBy,
    SheetVisibility,
    ClearApplyTo,
    ShapeTextHorizontalAlignment,
    ShapeTextVerticalAlignment,
    ChartLineStyle,
    ChartLegendPosition,
    RangeUnderlineStyle,
    ImageFittingMode,
    ErrorCodes,
  };

  const office = {
    actions: {
      associate(
        id: string,
        handler: (event?: { completed: () => void }) => void,
      ) {
        runtime.actions.set(id, handler);
      },
    },
    addin: { showAsTaskpane: () => Promise.resolve() },
    context: {
      requirements: {
        isSetSupported: (set: string, version: string) =>
          runtime.supported(set, version),
      },
      document: {
        addHandlerAsync: () => undefined,
        // Always synchronous here; the real call is async and can fail, which
        // is why the adapter reads the status rather than the value alone.
        getFilePropertiesAsync(
          callback: (result: {
            status: string;
            value: { url: string };
          }) => void,
        ) {
          callback({ status: "succeeded", value: { url: workbook.fileUrl } });
        },
      },
    },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    HostType: { Excel: "Excel", Word: "Word", PowerPoint: "PowerPoint" },
    EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
    onReady: (callback?: (info: { host: string }) => unknown) =>
      Promise.resolve(callback?.({ host: "Excel" })),
  };

  const scope = globalThis as unknown as Record<string, unknown>;
  scope.Excel = excel;
  scope.Office = office;

  const at = (address: string) => {
    const { sheet, rect } = resolve(workbook, address);
    return sheet.peek(rect.row, rect.col);
  };

  const helpers: FakeHelpers = {
    sheet(name) {
      const sheet = workbook.find(name);
      if (!sheet) throw new Error(`fake host: no sheet "${name}"`);
      return sheet;
    },
    addSheet: (name) => workbook.addSheet(name),
    deleteSheet: (name) => workbook.deleteSheet(name),
    addName(name, formula) {
      workbook.names.push({ name, formula, visible: true });
    },
    addComment(address, content, author, replies = []) {
      workbook.comments.push({
        address,
        content,
        author,
        replies: replies.map((reply) => ({
          content: reply.content,
          author: reply.author ?? author,
        })),
      });
    },
    // Excel rewrites a name's formula on its own when rows move under it; this
    // is how a test replays that without an insert-rows API.
    setNameFormula(name, formula) {
      const record = workbook.names.find((entry) => entry.name === name);
      if (!record) throw new Error(`fake host: no name "${name}"`);
      record.formula = formula;
    },
    breakName(name) {
      helpers.setNameFormula(name, "=#REF!");
    },
    select(address) {
      const { sheet, rect } = resolve(workbook, address);
      workbook.selection = { sheetId: sheet.id, rect };
      workbook.selectionAreas = [];
      workbook.activeSheetId = sheet.id;
      workbook.activeCell = null;
    },
    selectAreas(addresses) {
      const [first] = addresses;
      if (first === undefined) {
        throw new Error("fake host: a selection needs at least one area");
      }
      helpers.select(first);
      workbook.selectionAreas = addresses.map((address) => {
        const { sheet, rect } = resolve(workbook, address);
        return { sheetId: sheet.id, rect };
      });
    },
    setActiveCell(address) {
      const { sheet, rect } = resolve(workbook, address);
      workbook.activeCell = { sheetId: sheet.id, row: rect.row, col: rect.col };
    },
    seed(address, grid) {
      const { sheet, rect } = resolve(workbook, address);
      grid.forEach((row, r) => {
        row.forEach((entry, c) => {
          const cell = sheet.edit(rect.row + r, rect.col + c);
          if (entry !== null && typeof entry === "object") {
            if (entry.value !== undefined) cell.value = entry.value;
            if (entry.formula !== undefined) cell.formula = entry.formula;
            if (entry.r1c1 !== undefined) cell.formulaR1C1 = entry.r1c1;
            sheet.remember(
              rect.row + r,
              rect.col + c,
              cell.formula,
              cell.value,
            );
            return;
          }
          cell.value = entry;
          cell.formula = entry;
        });
      });
    },
    setNumberFormat(address, format) {
      const { sheet, rect } = resolve(workbook, address);
      for (let r = 0; r < rect.rowCount; r += 1) {
        for (let c = 0; c < rect.colCount; c += 1) {
          sheet.edit(rect.row + r, rect.col + c).numberFormat =
            runtime.format(format);
        }
      }
    },
    setFill(address, fill) {
      const { sheet, rect } = resolve(workbook, address);
      for (let r = 0; r < rect.rowCount; r += 1) {
        for (let c = 0; c < rect.colCount; c += 1) {
          Object.assign(sheet.edit(rect.row + r, rect.col + c).fill, fill);
        }
      }
    },
    setFont(address, font) {
      const { sheet, rect } = resolve(workbook, address);
      for (let r = 0; r < rect.rowCount; r += 1) {
        for (let c = 0; c < rect.colCount; c += 1) {
          Object.assign(sheet.edit(rect.row + r, rect.col + c).font, font);
        }
      }
    },
    // A style the workbook knows about; built-in ones are Excel's own, which no
    // add-in may delete.
    addStyle(name, builtIn = false) {
      workbook.styles.push({ name, builtIn });
    },
    setStyle(address, name) {
      const { sheet, rect } = resolve(workbook, address);
      for (let r = 0; r < rect.rowCount; r += 1) {
        for (let c = 0; c < rect.colCount; c += 1) {
          sheet.edit(rect.row + r, rect.col + c).style = name;
        }
      }
    },
    cell: (address) => clone(at(address)),
    fill: (address) => clone(at(address).fill),
    font: (address) => clone(at(address).font),
    border: (address, edge) => clone(at(address).borders[edge]),
    numberFormat: (address) => at(address).numberFormat,
    value: (address) => at(address).value,
    formula: (address) => at(address).formula,
    cellMap(sheetName) {
      const sheet = helpers.sheet(sheetName);
      const out: Record<string, FakeCell> = {};
      // Untouched and reverted-to-default cells are the same observable state.
      for (const [key, cell] of sheet.cells) {
        if (!isDefaultCell(cell)) out[key] = clone(cell);
      }
      return out;
    },
    rowHeight: (sheetName, row) =>
      helpers.sheet(sheetName).rowHeights.get(row) ?? DEFAULT_ROW_HEIGHT,
    columnWidth: (sheetName, col) =>
      helpers.sheet(sheetName).columnWidths.get(col) ?? DEFAULT_COLUMN_WIDTH,
    setPrecedents(address, config) {
      workbook.precedents.set(address, config);
    },
    setDependents(address, config) {
      workbook.dependents.set(address, config);
    },
    addChart(sheetName, chart = {}) {
      const record = newChart(
        sheetName,
        `Chart ${workbook.charts.length + 1}`,
        chart,
      );
      workbook.charts.push(record);
      return record;
    },
    moveChart(name, toSheet) {
      const record = workbook.charts.find((chart) => chart.name === name);
      if (!record) throw new Error(`fake host: no chart "${name}"`);
      if (!workbook.find(toSheet)) {
        throw new Error(`fake host: no sheet "${toSheet}"`);
      }
      record.sheetName = toSheet;
    },
    setActiveChart(chart) {
      workbook.activeChart = chart;
    },
    setting: (key) => workbook.settings.get(key) ?? null,
    setSetting(key, value) {
      workbook.settings.set(key, value);
    },
    setSupported(check) {
      runtime.supported = check;
    },
    failNextSync(error) {
      runtime.failSync =
        error ?? hostError(ErrorCodes.generalException, "The sync failed.");
    },
    failNextImage(error) {
      runtime.failImage =
        error ??
        hostError(ErrorCodes.generalException, "The image failed to render.");
    },
    changeHandlerCount: () => runtime.changeHandlers.length,
    syncCount: () => runtime.syncs,
    async fireChanged(sheetIdOrName, address) {
      const sheet = workbook.find(sheetIdOrName);
      if (!sheet) throw new Error(`fake host: no sheet "${sheetIdOrName}"`);
      const areas = address.split(",").map((area) => formatA1(parseA1(area)));
      const local = areas[0] ?? "";
      const args = {
        address: areas
          .map((area) => `${quoteSheet(sheet.name)}!${area}`)
          .join(","),
        worksheetId: sheet.id,
        changeType: "RangeEdited",
        source: "Local",
        type: "WorksheetChanged",
        getRange: (context: FakeContext) =>
          context.workbook.worksheets.getItem(sheet.id).getRange(local),
        getRangeOrNullObject: (context: FakeContext) =>
          context.workbook.worksheets.getItem(sheet.id).getRange(local),
      };
      for (const registration of [...runtime.changeHandlers]) {
        await registration.handler(args);
      }
    },
    actions: () => runtime.actions,
  };

  return { workbook, helpers };
}

export function uninstallFakeHost(): void {
  const scope = globalThis as unknown as Record<string, unknown>;
  delete scope.Excel;
  delete scope.Office;
}
