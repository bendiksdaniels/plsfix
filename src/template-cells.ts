// The vocabulary a template grid is written in: the cell kinds and number
// formats, the `{r+dr,c+dc}` placeholder grammar that names a neighbour from
// the block's own top-left corner, and one constructor per look. Owns the
// grammar; the invariant is that a placeholder is an offset, never an address,
// so the same grid resolves anywhere on any sheet.
import { cellAddress } from "./find";
import { type NumberFormatName } from "./excel/shared";

export type TemplateCellKind =
  "title" | "header" | "label" | "input" | "formula" | "result";

/** `percent` is already one of the number-format names; `multiple` is not. */
export type TemplateFormat = NumberFormatName | "multiple";

export interface TemplateCell {
  /** A literal, for a label or a hardcoded input. Never set beside `f`. */
  v?: number | string;
  /** A formula with `{r+dr,c+dc}` placeholders, `=` and all. */
  f?: string;
  kind: TemplateCellKind;
  format?: TemplateFormat;
}

/** Zero-based, either a sheet coordinate or an offset from the block's origin. */
export interface CellRef {
  row: number;
  column: number;
}

const PLACEHOLDER = /\{r([+-]\d+)?,c([+-]\d+)?\}/g;

/** Which block cells a formula names, in the order it names them. */
export function placeholderRefs(formula: string): CellRef[] {
  return [...formula.matchAll(PLACEHOLDER)].map((match) => ({
    row: Number(match[1] ?? 0),
    column: Number(match[2] ?? 0),
  }));
}

/** The same formula with every placeholder turned into an A1 address. */
export function resolveFormula(formula: string, origin: CellRef): string {
  return formula.replace(
    PLACEHOLDER,
    (_match: string, row?: string, column?: string) =>
      cellAddress(
        origin.row + Number(row ?? 0),
        origin.column + Number(column ?? 0),
      ),
  );
}

// ---------------------------------------------------------------------------
// Cell constructors: one per look, so a grid literal reads as the block does.
// ---------------------------------------------------------------------------

export function textCell(v: string, kind: TemplateCellKind): TemplateCell {
  return { v, kind };
}

export function numberCell(
  v: number,
  kind: TemplateCellKind,
  format: TemplateFormat,
): TemplateCell {
  return { v, kind, format };
}

export function formulaCell(
  f: string,
  kind: TemplateCellKind,
  format: TemplateFormat,
): TemplateCell {
  return { f, kind, format };
}

export const title = (v: string): TemplateCell => textCell(v, "title");
export const head = (v: string): TemplateCell => textCell(v, "header");
export const label = (v: string): TemplateCell => textCell(v, "label");
export const result = (v: string): TemplateCell => textCell(v, "result");
export const blank = (kind: TemplateCellKind = "label"): TemplateCell => ({
  kind,
});

export const input = (v: number, format: TemplateFormat): TemplateCell =>
  numberCell(v, "input", format);
export const calc = (f: string, format: TemplateFormat): TemplateCell =>
  formulaCell(f, "formula", format);
export const total = (f: string, format: TemplateFormat): TemplateCell =>
  formulaCell(f, "result", format);
export const axis = (f: string): TemplateCell =>
  formulaCell(f, "header", "percent");

export function blankRow(
  width: number,
  kind: TemplateCellKind = "label",
): TemplateCell[] {
  return Array.from({ length: width }, () => blank(kind));
}

/** Fills a short row out to the block width with empty cells of one kind. */
export function pad(
  row: TemplateCell[],
  width: number,
  kind: TemplateCellKind = "label",
): TemplateCell[] {
  return [...row, ...blankRow(width - row.length, kind)];
}

export interface Template {
  id: string;
  name: string;
  description: string;
  rows: number;
  cols: number;
  cells: TemplateCell[][];
}

export function build(
  id: string,
  name: string,
  description: string,
  cells: TemplateCell[][],
): Template {
  return {
    id,
    name,
    description,
    rows: cells.length,
    cols: cells[0]?.length ?? 0,
    cells,
  };
}
