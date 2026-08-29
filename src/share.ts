// Prepare for sharing: what a model still carries when it leaves the desk, and
// how that run reads back in one line. Pure - the adapter in src/excel/share.ts
// does the reading and the A1 reset; this file only decides what counts as
// something the reader would find and how it is worded.

import { classifyCell } from "./classify";
import { type CellValue, isFormula } from "./model";

export type ShareIssueKind =
  | "hiddenSheet"
  | "externalLink"
  | "brokenName"
  | "autocolorOnEdit"
  | "skippedSheet";

export interface ShareIssue {
  kind: ShareIssueKind;
  label: string;
}

export interface ShareSheet {
  name: string;
  visibility: string;
}

export interface ShareLink {
  sheet: string;
  address: string;
  formula: string;
}

export interface ShareScan {
  sheets: ShareSheet[];
  externalLinks: ShareLink[];
  brokenNames: string[];
  // Sheets whose used range was too large to read, so the report can say the
  // answer is incomplete rather than quietly leaving them out.
  skippedSheets: string[];
  autocolorOnEdit: boolean;
}

export const AUTOCOLOR_LABEL = "Autocolor recolors cells as they are edited";

// A reference to another workbook is bracketed - ='[Budget.xlsx]Model'!$B$4 -
// which is exactly the class the color key already paints as external, so the
// scan asks the same classifier instead of inventing a second rule. Both
// brackets are required: a lone "[" is a half-typed formula, not a link.
export function isExternalFormula(formula: CellValue): formula is string {
  if (!isFormula(formula)) return false;
  return classifyCell(formula, null) === "external" && formula.includes("]");
}

function isVisible(sheet: ShareSheet): boolean {
  return sheet.visibility.toLowerCase() === "visible";
}

// Very hidden is set outside Excel's UI and cannot be shown from one, so the
// reader is told which kind of hidden they are looking at.
function hiddenLabel(sheet: ShareSheet): string {
  const veryHidden = sheet.visibility.toLowerCase() === "veryhidden";
  return veryHidden ? `${sheet.name} (very hidden)` : sheet.name;
}

function linkLabel(link: ShareLink): string {
  return `${link.sheet}!${link.address}: ${link.formula}`;
}

// Workbook order: the sheets a reader would find, then what the formulas point
// at, then the names, then what could not be read, then our own live handler.
export function shareReport(scan: ShareScan): ShareIssue[] {
  const hidden = scan.sheets.filter((sheet) => !isVisible(sheet));
  return [
    ...hidden.map((sheet): ShareIssue => ({
      kind: "hiddenSheet",
      label: hiddenLabel(sheet),
    })),
    ...scan.externalLinks.map((link): ShareIssue => ({
      kind: "externalLink",
      label: linkLabel(link),
    })),
    ...scan.brokenNames.map((name): ShareIssue => ({
      kind: "brokenName",
      label: name,
    })),
    ...scan.skippedSheets.map((name): ShareIssue => ({
      kind: "skippedSheet",
      label: name,
    })),
    ...(scan.autocolorOnEdit
      ? [{ kind: "autocolorOnEdit" as const, label: AUTOCOLOR_LABEL }]
      : []),
  ];
}

const COUNTED: { kind: ShareIssueKind; one: string; many: string }[] = [
  { kind: "hiddenSheet", one: "hidden sheet", many: "hidden sheets" },
  { kind: "externalLink", one: "external link", many: "external links" },
  { kind: "brokenName", one: "broken name", many: "broken names" },
  {
    kind: "skippedSheet",
    one: "sheet too large to scan",
    many: "sheets too large to scan",
  },
];

function countedParts(report: ShareIssue[]): string[] {
  const parts: string[] = [];
  for (const { kind, one, many } of COUNTED) {
    const count = report.filter((issue) => issue.kind === kind).length;
    if (count > 0) parts.push(`${count} ${count === 1 ? one : many}`);
  }
  return parts;
}

// One line for the toast and the pane: what the run did, then what it found and
// deliberately did not touch. Autocolor is a switch rather than a count, so it
// gets its own clause instead of being counted alongside the sheets.
export function summarizeShare(
  report: ShareIssue[],
  touchedSheets: number,
): string {
  const sheets = touchedSheets === 1 ? "sheet" : "sheets";
  const clauses = [`${touchedSheets} ${sheets} reset to A1`];

  const counted = countedParts(report);
  if (counted.length > 0) clauses.push(`${counted.join(", ")} left`);
  if (report.some((issue) => issue.kind === "autocolorOnEdit")) {
    clauses.push("autocolor on edit is still on");
  }
  if (clauses.length === 1) clauses.push("nothing else to fix");

  return clauses.join("; ");
}
