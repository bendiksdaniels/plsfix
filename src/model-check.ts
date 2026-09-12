// Model check: the one pass a reviewer makes over a workbook, as rules and
// wording. Owns what counts as a finding, the order the findings read in, the
// cap on the list, and the two strings the pane shows: the toast line and the
// plain-text report the Copy button hands over. Pure, in and out: the reading
// lives in src/excel/model-check.ts.

import { hasHardcodedNumber } from "./classify";
import { isFormula } from "./model";

export type CheckKind =
  | "formulaError"
  | "hardcodeInFormula"
  | "inconsistentFormula"
  | "volatileFormula"
  | "brokenName"
  | "unusedStyle"
  | "hiddenSheet"
  | "externalLink";

export interface Finding {
  kind: CheckKind;
  // Null for a finding the whole file carries rather than one sheet: a broken
  // name, a style nothing wears.
  sheet: string | null;
  // Null where there is no cell to jump to.
  ref: string | null;
  count: number;
  note: string;
}

export interface ModelCheckReport {
  findings: Finding[];
  scanned: { sheets: number; cells: number };
  // Sheets whose used range was too large to read, by name: the report says
  // the answer is incomplete rather than quietly leaving them out.
  skipped: string[];
  // The workbook had more findings than the cap: the list is the first page of
  // them, and every line the pane shows says so.
  truncated: boolean;
}

export const KIND_LABELS: Record<CheckKind, string> = {
  formulaError: "Formula error",
  hardcodeInFormula: "Hardcode in formula",
  inconsistentFormula: "Inconsistent formula",
  volatileFormula: "Volatile function",
  brokenName: "Broken name",
  unusedStyle: "Unused style",
  hiddenSheet: "Hidden sheet",
  externalLink: "External link",
};

// A list is read, not scrolled: past this the pane is a worse tool than a
// filtered sheet, and every extra line costs the reviewer attention.
export const MODEL_CHECK_FINDING_CAP = 500;

// The volatile functions a reviewer flags: each one recalculates on every
// change anywhere in the workbook, and OFFSET and INDIRECT also hide their
// precedents from the trace tools.
const VOLATILE = new Set([
  "OFFSET",
  "INDIRECT",
  "NOW",
  "TODAY",
  "RAND",
  "RANDBETWEEN",
  "CELL",
  "INFO",
]);

// A function call is a word token followed by an opening bracket. The token is
// captured whole, so RANDBETWEEN reads as itself rather than as RAND, and the
// _xlfn. prefix Excel writes on a function it did not resolve is cut off before
// the name is matched.
const CALL = /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;

/**
 * The first volatile function called in a formula, uppercased, else null.
 * Case-insensitive: Excel stores what the modeller typed.
 */
export function volatileIn(formula: string): string | null {
  CALL.lastIndex = 0;
  let match = CALL.exec(formula);
  while (match !== null) {
    const token = (match[1] ?? "").split(".").pop() ?? "";
    const name = token.toUpperCase();
    if (VOLATILE.has(name)) {
      CALL.lastIndex = 0;
      return name;
    }
    match = CALL.exec(formula);
  }
  return null;
}

/**
 * A number typed into a formula. The same rule the color key paints as a
 * partial input, asked here of every formula rather than only of the ones that
 * point at this sheet.
 */
export function hardcodeIn(formula: string): boolean {
  return isFormula(formula) && hasHardcodedNumber(formula);
}

/** Where a finding sits, as the report and the pane both write it. */
export function findingLocation(finding: Finding): string {
  if (finding.sheet === null) return "Workbook";
  return finding.ref === null
    ? finding.sheet
    : `${finding.sheet}!${finding.ref}`;
}

const A1 = /^([A-Za-z]+)(\d+)$/;

// Reading order, not alphabetical: "A10" sorts after "A9" the way the grid
// does. A finding with no cell sorts ahead of the first one, because it is
// about the sheet rather than about a place on it.
function cellOrder(ref: string | null): [number, number] {
  const match = ref === null ? null : A1.exec(ref);
  if (match === null) return [-1, -1];
  let column = 0;
  for (const letter of (match[1] ?? "").toUpperCase()) {
    column = column * 26 + (letter.charCodeAt(0) - 64);
  }
  return [Number(match[2]), column];
}

// Workbook order, and the file's own findings after every sheet: a broken name
// or an unused style belongs to the workbook, not to whatever sheet happens to
// sort last.
function sheetOrderOf(sheet: string | null, order: string[]): number {
  if (sheet === null) return order.length;
  const index = order.indexOf(sheet);
  return index < 0 ? order.length : index;
}

// Errors first within a sheet: a cell that does not compute outranks a cell
// that computes the wrong way.
function compare(left: Finding, right: Finding, order: string[]): number {
  const bySheet =
    sheetOrderOf(left.sheet, order) - sheetOrderOf(right.sheet, order);
  if (bySheet !== 0) return bySheet;

  const errorFirst =
    Number(left.kind !== "formulaError") -
    Number(right.kind !== "formulaError");
  if (errorFirst !== 0) return errorFirst;

  const [leftRow, leftColumn] = cellOrder(left.ref);
  const [rightRow, rightColumn] = cellOrder(right.ref);
  return leftRow - rightRow || leftColumn - rightColumn;
}

export interface ReportInput {
  findings: Finding[];
  /** Sheet names in workbook order, which is the order findings read in. */
  sheetOrder: string[];
  scanned: { sheets: number; cells: number };
  skipped: string[];
}

/** Orders the findings, caps the list and says whether the cap was reached. */
export function buildReport(input: ReportInput): ModelCheckReport {
  const ordered = [...input.findings].sort((left, right) =>
    compare(left, right, input.sheetOrder),
  );
  return {
    findings: ordered.slice(0, MODEL_CHECK_FINDING_CAP),
    scanned: input.scanned,
    skipped: input.skipped,
    truncated: ordered.length > MODEL_CHECK_FINDING_CAP,
  };
}

/** "1 cell" / "2 cells": the count and its noun, for any line that counts. */
export function plural(count: number, one: string): string {
  return `${String(count)} ${count === 1 ? one : `${one}s`}`;
}

/**
 * The one line the toast and the pane hint show: what was found, over how much
 * of the workbook, and what the answer is missing.
 */
export function summarize(report: ModelCheckReport): string {
  const sheets = plural(report.scanned.sheets, "sheet");
  const head =
    report.findings.length === 0
      ? `Model check: nothing to flag on ${sheets}`
      : `Model check: ${plural(report.findings.length, "finding")} on ${sheets}`;

  const notes: string[] = [];
  if (report.truncated) {
    notes.push(`list capped at ${String(MODEL_CHECK_FINDING_CAP)}`);
  }
  if (report.skipped.length > 0) {
    notes.push(`${plural(report.skipped.length, "sheet")} skipped`);
  }
  return notes.length === 0 ? head : `${head} (${notes.join(", ")})`;
}

/** The plain-text report the Copy button copies: one line per finding. */
export function reportText(report: ModelCheckReport): string {
  return report.findings
    .map((finding) =>
      [findingLocation(finding), KIND_LABELS[finding.kind], finding.note].join(
        "  ",
      ),
    )
    .join("\n");
}
