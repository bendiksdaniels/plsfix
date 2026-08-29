// The table a Table shape hands out: PowerPoint.Table and its cells over the
// FakeTable the deck stores on the shape. Owns the cell grid and the writes an
// add-in makes on it - text, font, fill and alignment. Invariant: every object
// addresses its cell by (row, column) and resolves the shape again on each
// access, so a table deleted under it is seen as gone.

import type { FakePresentation, FakePptShape } from "./model";
import { Loadable } from "./strict";

export interface FakeTableCell {
  text: string;
  font: {
    bold: boolean | null;
    italic: boolean | null;
    color: string | null;
    size: number | null;
  };
  fill: { color: string | null };
  horizontalAlignment: string | null;
}

export interface FakeTable {
  rowCount: number;
  columnCount: number;
  cells: FakeTableCell[][];
}

// A fresh grid: PowerPoint fills the cells addTable was given values for and
// leaves the rest empty, with every format inherited from the table style.
export function newFakeTable(
  rowCount: number,
  columnCount: number,
  values: string[][] = [],
): FakeTable {
  const cells = Array.from({ length: rowCount }, (_row, r) =>
    Array.from({ length: columnCount }, (_column, c): FakeTableCell => ({
      text: values[r]?.[c] ?? "",
      font: { bold: null, italic: null, color: null, size: null },
      fill: { color: null },
      horizontalAlignment: null,
    })),
  );
  return { rowCount, columnCount, cells };
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
  getCellOrNullObject(rowIndex: number, columnIndex: number): TableCellProxy {
    return new TableCellProxy(this.deck, this.shapeId, rowIndex, columnIndex);
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
  setSolidColor(color: string): void {
    this.cell.fill.color = color;
  }
  clear(): void {
    this.cell.fill.color = null;
  }
}
