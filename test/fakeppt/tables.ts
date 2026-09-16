// The table a Table shape hands out: PowerPoint.Table, its style settings and
// its cells, over the FakeTable the deck stores on the shape. Owns the cell
// grid, the style settings and the writes an add-in makes on either - text,
// font, fill, alignment and the header band. Invariant: every object
// addresses its cell by (row, column) and resolves the shape again on each
// access, so a table deleted under it is seen as gone.

import type { FakePresentation, FakePptShape } from "./model";
import { Loadable } from "./strict";

// PowerPoint's own default for a table added through shapes.addTable: Medium
// Style 2 Accent 1, header row highlighted, rows banded, nothing else. A
// literal, not an import from src/ppt/table-style.ts: this fixture answers
// for what the HOST hands back, whatever pls,fix later does with it.
const DEFAULT_TABLE_STYLE = "MediumStyle2Accent1";

// `color` and `cleared` are the state PowerPoint keeps; `type` is derived,
// never stored twice, the same way format.ts's ShapeFillProxy derives a
// shape's: Solid once a colour is set, NoFill once clear() ran with none,
// Inherited - the table style's own band still showing through - for a cell
// nothing has written yet.
export class FakeTableCellFill {
  color: string | null = null;
  cleared = false;
  get type(): string {
    if (this.color !== null) return "Solid";
    return this.cleared ? "NoFill" : "Inherited";
  }
}

export interface FakeTableCell {
  text: string;
  font: {
    bold: boolean | null;
    italic: boolean | null;
    color: string | null;
    size: number | null;
  };
  fill: FakeTableCellFill;
  horizontalAlignment: string | null;
}

export interface FakeTableStyleSettings {
  style: string;
  isFirstRowHighlighted: boolean;
  areRowsBanded: boolean;
  isFirstColumnHighlighted: boolean;
  isLastRowHighlighted: boolean;
  isLastColumnHighlighted: boolean;
  areColumnsBanded: boolean;
}

// The one property pls,fix ever sets on TableAddOptions.uniformCellProperties:
// real PowerPoint.TableCellProperties carries far more (fill, borders,
// alignment...), but the fake only needs to prove this one travels.
export interface TableUniformCellProperties {
  font?: { size?: number };
}

export interface FakeTable {
  /** What addTable was asked for per column, null where it was left to PowerPoint. */
  columnWidths: (number | null)[];
  rowCount: number;
  columnCount: number;
  cells: FakeTableCell[][];
  styleSettings: FakeTableStyleSettings;
  /** What addTable's own uniformCellProperties carried, exactly as given;
   * null when the add named none - a test's proof that the size PowerPoint
   * creates every row at came from the call itself, not a coincidence. */
  uniformCellProperties: TableUniformCellProperties | null;
}

function defaultStyleSettings(): FakeTableStyleSettings {
  return {
    style: DEFAULT_TABLE_STYLE,
    isFirstRowHighlighted: true,
    areRowsBanded: true,
    isFirstColumnHighlighted: false,
    isLastRowHighlighted: false,
    isLastColumnHighlighted: false,
    areColumnsBanded: false,
  };
}

// A fresh grid: PowerPoint fills the cells addTable was given values for and
// leaves the rest empty, with every format inherited from the table style -
// except the font size, which uniformCellProperties sets on every cell right
// away, the same as the real host does.
export function newFakeTable(
  rowCount: number,
  columnCount: number,
  values: string[][] = [],
  uniformCellProperties: TableUniformCellProperties | null = null,
): FakeTable {
  const uniformSize = uniformCellProperties?.font?.size ?? null;
  const cells = Array.from({ length: rowCount }, (_row, r) =>
    Array.from({ length: columnCount }, (_column, c): FakeTableCell => ({
      text: values[r]?.[c] ?? "",
      font: { bold: null, italic: null, color: null, size: uniformSize },
      fill: new FakeTableCellFill(),
      horizontalAlignment: null,
    })),
  );
  return {
    columnWidths: [],
    rowCount,
    columnCount,
    cells,
    styleSettings: defaultStyleSettings(),
    uniformCellProperties,
  };
}

// A test's one-shot override of the style a table about to be added will
// start with, keyed by the shape id addTable is about to assign - ids are
// sequential per deck, so a fresh presentation's first table is always
// "shape-1". Consumed the moment that shape is created; no production caller
// ever reaches this, since pls,fix only ever reads a style PowerPoint already
// handed back.
const pendingStyles = new WeakMap<FakePresentation, Map<string, string>>();

export function setPendingTableStyle(
  deck: FakePresentation,
  shapeId: string,
  style: string,
): void {
  const styles = pendingStyles.get(deck) ?? new Map<string, string>();
  styles.set(shapeId, style);
  pendingStyles.set(deck, styles);
}

function takePendingTableStyle(
  deck: FakePresentation,
  shapeId: string,
): string | null {
  const styles = pendingStyles.get(deck);
  const style = styles?.get(shapeId) ?? null;
  styles?.delete(shapeId);
  return style;
}

// Applied right after addTable creates the shape: a no-op for every table
// that is not one a test pre-registered.
export function applyPendingTableStyle(
  deck: FakePresentation,
  shape: FakePptShape,
): void {
  const forced = takePendingTableStyle(deck, shape.id);
  if (forced !== null && shape.table) shape.table.styleSettings.style = forced;
}

function fail(message: string, code: string): Error {
  return Object.assign(new Error(`${code}: ${message}`), { code });
}

abstract class TableBound extends Loadable {
  constructor(
    protected deck: FakePresentation,
    protected shapeId: string,
  ) {
    super();
  }

  protected table(): FakeTable {
    const shape: FakePptShape = this.deck.findShape(this.shapeId).shape;
    if (!shape.table) {
      throw fail(`shape "${this.shapeId}" is no table`, "GeneralException");
    }
    return shape.table;
  }
}

export class TableProxy extends TableBound {
  get rowCount(): number {
    return this.table().rowCount;
  }
  get columnCount(): number {
    return this.table().columnCount;
  }
  get styleSettings(): TableStyleSettingsProxy {
    return new TableStyleSettingsProxy(this.deck, this.shapeId);
  }
  getCellOrNullObject(rowIndex: number, columnIndex: number): TableCellProxy {
    return new TableCellProxy(this.deck, this.shapeId, rowIndex, columnIndex);
  }
}

class TableStyleSettingsProxy extends TableBound {
  get style(): string {
    return this.table().styleSettings.style;
  }
  set style(value: string) {
    this.table().styleSettings.style = value;
  }
  get isFirstRowHighlighted(): boolean {
    return this.table().styleSettings.isFirstRowHighlighted;
  }
  set isFirstRowHighlighted(value: boolean) {
    this.table().styleSettings.isFirstRowHighlighted = value;
  }
  get areRowsBanded(): boolean {
    return this.table().styleSettings.areRowsBanded;
  }
  set areRowsBanded(value: boolean) {
    this.table().styleSettings.areRowsBanded = value;
  }
  get isFirstColumnHighlighted(): boolean {
    return this.table().styleSettings.isFirstColumnHighlighted;
  }
  set isFirstColumnHighlighted(value: boolean) {
    this.table().styleSettings.isFirstColumnHighlighted = value;
  }
  get isLastRowHighlighted(): boolean {
    return this.table().styleSettings.isLastRowHighlighted;
  }
  set isLastRowHighlighted(value: boolean) {
    this.table().styleSettings.isLastRowHighlighted = value;
  }
  get isLastColumnHighlighted(): boolean {
    return this.table().styleSettings.isLastColumnHighlighted;
  }
  set isLastColumnHighlighted(value: boolean) {
    this.table().styleSettings.isLastColumnHighlighted = value;
  }
  get areColumnsBanded(): boolean {
    return this.table().styleSettings.areColumnsBanded;
  }
  set areColumnsBanded(value: boolean) {
    this.table().styleSettings.areColumnsBanded = value;
  }
}

// A cell outside the grid is refused rather than handed back as a null object:
// the add-in only ever addresses the cells its own payload describes, and a
// silent no-op would hide the day it stops doing that.
class TableCellProxy extends TableBound {
  constructor(
    deck: FakePresentation,
    shapeId: string,
    private row: number,
    private column: number,
  ) {
    super(deck, shapeId);
  }

  get text(): string {
    return this.cell().text;
  }
  set text(value: string) {
    this.cell().text = value;
  }
  get horizontalAlignment(): string | null {
    return this.cell().horizontalAlignment;
  }
  set horizontalAlignment(value: string | null) {
    this.cell().horizontalAlignment = value;
  }
  get font(): TableCellFontProxy {
    return new TableCellFontProxy(this.cell());
  }
  get fill(): TableCellFillProxy {
    return new TableCellFillProxy(this.cell());
  }

  private cell(): FakeTableCell {
    const cell = this.table().cells[this.row]?.[this.column];
    if (!cell) {
      throw fail(
        `no cell ${String(this.row)},${String(this.column)}`,
        "ItemNotFound",
      );
    }
    return cell;
  }
}

class TableCellFontProxy extends Loadable {
  constructor(private cell: FakeTableCell) {
    super();
  }
  get bold(): boolean | null {
    return this.cell.font.bold;
  }
  set bold(value: boolean | null) {
    this.cell.font.bold = value;
  }
  get italic(): boolean | null {
    return this.cell.font.italic;
  }
  set italic(value: boolean | null) {
    this.cell.font.italic = value;
  }
  get color(): string | null {
    return this.cell.font.color;
  }
  set color(value: string | null) {
    this.cell.font.color = value;
  }
  get size(): number | null {
    return this.cell.font.size;
  }
  set size(value: number | null) {
    this.cell.font.size = value;
  }
}

class TableCellFillProxy extends Loadable {
  constructor(private cell: FakeTableCell) {
    super();
  }
  get type(): string {
    return this.cell.fill.type;
  }
  setSolidColor(color: string): void {
    this.cell.fill.color = color;
    this.cell.fill.cleared = false;
  }
  clear(): void {
    this.cell.fill.color = null;
    this.cell.fill.cleared = true;
  }
}
