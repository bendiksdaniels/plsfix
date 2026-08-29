// Prepare for sharing: what a model still carries when it leaves the desk, and
// how that run reads back in one line. Pure - the adapter in src/excel/share.ts
// does the reading and the A1 reset; this file only decides what counts as
// something the reader would find and how it is worded.

import { classifyCell } from "./classify";
import { type CellValue, isFormula } from "./model";

export type ShareIssueKind =
  | "hiddenSheet"
  | "externalLink"
  | "addinFormula"
  | "brokenName"
  | "skippedSheet"
  | "overlayPainted"
  | "linkTokens"
  | "autocolorOnEdit";

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
  // Cells only this add-in can evaluate: a reader without it gets #NAME?.
  addinFormulas: ShareLink[];
  brokenNames: string[];
  // Sheets whose used range was too large to read, so the report can say the
  // answer is incomplete rather than quietly leaving them out.
  skippedSheets: string[];
  // The overlays still holding fills, by the name the pane calls them: paint
  // travels with the file, and the reader has no add-in to take it off.
  overlaysPainted: string[];
  // The workbook carries a link registry. Each entry holds the token that
  // decrypts that link's picture on the relay, so the file is the key.
  linkTokens: boolean;
  autocolorOnEdit: boolean;
}

export const AUTOCOLOR_LABEL = "Autocolor recolors cells as they are edited";
export const LINK_TOKENS_LABEL =
  "Link registry present: break the links before sharing this file outside the team";

// A reference to another workbook is bracketed - ='[Budget.xlsx]Model'!$B$4 -
// which is exactly the class the color key already paints as external, so the
// scan asks the same classifier instead of inventing a second rule. Both
// brackets are required: a lone "[" is a half-typed formula, not a link.
export function isExternalFormula(formula: CellValue): formula is string {
  if (!isFormula(formula)) return false;
  return classifyCell(formula, null) === "external" && formula.includes("]");
}

// =PLSFIX.ROUND and =PLSFIX.ROUNDSUM are this add-in's own functions. Opened without
// it, Excel keeps them as _xlfn.PLSFIX.ROUND and every one of those cells reads
// #NAME?, so a model full of them is only a model on a machine that has the
// add-in. The namespace is matched either way round: the stored formula carries
// the _xlfn. prefix once the workbook has been opened somewhere without us.
export function isAddinFormula(formula: CellValue): formula is string {
  if (!isFormula(formula)) return false;
  return formula.toUpperCase().includes("PLSFIX.");
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

function overlayLabel(name: string): string {
  return `${name} is still painted over cells; the reader has no add-in to clear it`;
}

// What the file itself carries once it leaves: the paint on the cells, the
// tokens in the settings and our own live handler. Reported, never undone -
// this pass tidies, it does not delete.
function stateIssues(scan: ShareScan): ShareIssue[] {
  return [
    ...scan.overlaysPainted.map((name): ShareIssue => ({
      kind: "overlayPainted",
      label: overlayLabel(name),
    })),
    ...(scan.linkTokens
      ? [{ kind: "linkTokens" as const, label: LINK_TOKENS_LABEL }]
      : []),
    ...(scan.autocolorOnEdit
      ? [{ kind: "autocolorOnEdit" as const, label: AUTOCOLOR_LABEL }]
      : []),
  ];
}

// Workbook order: the sheets a reader would find, then what the formulas point
// at and which of them need us, then the names, then what could not be read,
// then the state the file is carrying.
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
    ...scan.addinFormulas.map((link): ShareIssue => ({
      kind: "addinFormula",
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
    ...stateIssues(scan),
  ];
}

const COUNTED: { kind: ShareIssueKind; one: string; many: string }[] = [
  { kind: "hiddenSheet", one: "hidden sheet", many: "hidden sheets" },
  { kind: "externalLink", one: "external link", many: "external links" },
  {
    kind: "addinFormula",
    one: "cell needing the add-in",
    many: "cells needing the add-in",
  },
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
// deliberately did not touch. Autocolor, the overlays and the link registry are
// switches rather than counts, so each gets its own clause instead of being
// counted alongside the sheets.
export function summarizeShare(
  report: ShareIssue[],
  touchedSheets: number,
): string {
  const sheets = touchedSheets === 1 ? "sheet" : "sheets";
  const clauses = [`${touchedSheets} ${sheets} reset to A1`];
  const has = (kind: ShareIssueKind): boolean =>
    report.some((issue) => issue.kind === kind);

  const counted = countedParts(report);
  if (counted.length > 0) clauses.push(`${counted.join(", ")} left`);
  if (has("autocolorOnEdit")) clauses.push("autocolor on edit is still on");
  if (has("overlayPainted")) clauses.push("an overlay is still painted");
  if (has("linkTokens")) clauses.push("link tokens travel with the file");
  if (clauses.length === 1) clauses.push("nothing else to fix");

  return clauses.join("; ");
}
