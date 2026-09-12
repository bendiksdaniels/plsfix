// Classifies a cell by its formula and value into the pane's colour key:
// blank, input, formula, crossSheet or external, plus "partial" for a formula
// that also hardcodes a number. Pure, no Office.js. Invariant: a "[" reads as
// another workbook only where a sheet name and "!" actually follow it.
import { type CellValue, isFormula } from "./model";

export type CellClass =
  "blank" | "input" | "formula" | "crossSheet" | "external" | "partial";

// Cell references ($B$12), function names (LOG10) and range operators all read as
// word tokens; whatever digits survive their removal were typed by the modeller.
const WORD_TOKEN = /[A-Za-z_$][A-Za-z0-9_.$]*/g;
const DIGIT = /[0-9]/;

// Quoted text can hold anything ("wow!", "[note]", "2026"), so it goes first.
function stripStringLiterals(formula: string): string {
  let stripped = "";
  let inLiteral = false;

  for (let index = 0; index < formula.length; index += 1) {
    const char = formula[index];
    if (!inLiteral) {
      if (char === '"') inLiteral = true;
      else stripped += char;
      continue;
    }
    if (char !== '"') continue;
    // A doubled quote is an escaped quote, not the end of the literal.
    if (formula[index + 1] === '"') index += 1;
    else inLiteral = false;
  }

  return stripped;
}

// The characters a table's own name is made of: a "[" behind one of them opens
// a structured reference (Table1[Revenue]), never a workbook reference.
const NAME_CHAR = /[A-Za-z0-9_.$]/;

// From the opening bracket to its match, so the "[" and "," inside
// Sales[[#Headers],[Amount]] cannot be read as references of their own.
function matchBracket(body: string, open: number): number {
  let depth = 0;
  for (let index = open; index < body.length; index += 1) {
    if (body[index] === "[") depth += 1;
    else if (body[index] === "]") {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return body.length;
}

// What a sheet name behind a workbook bracket is made of, quotes and spaces
// included: '[Budget.xlsx]Model plan'!$B$4.
const SHEET_CHAR = /[A-Za-z0-9_.$' ]/;

// A workbook bracket is always followed by its sheet and a "!". A structured
// reference outside a table name ([@Amount], [Amount], [@[Unit price]]) is
// followed by an operator, a bracket or nothing, never by a sheet name.
function sheetFollows(body: string, close: number): boolean {
  let index = close + 1;
  while (SHEET_CHAR.test(body[index] ?? "")) index += 1;
  return body[index] === "!";
}

// A bracket is another workbook - '[Budget.xlsx]Model'!$B$4, [1]Sheet1!A1 - only
// where a reference may start and a sheet name follows it. Behind a name it is
// this workbook's own table, which is neither a link nor anything to warn about.
function hasWorkbookReference(body: string): boolean {
  for (let index = 0; index < body.length; index += 1) {
    if (body[index] !== "[") continue;
    const close = matchBracket(body, index);
    if (!NAME_CHAR.test(body[index - 1] ?? "") && sheetFollows(body, close)) {
      return true;
    }
    index = close;
  }
  return false;
}

// A constant in a comparison counts as a hardcode: the threshold is an
// assumption that belongs in its own cell. Exported because the model check
// asks the same question of formulas the color key has already answered
// "crossSheet" or "external" for, and one rule must have one home.
export function hasHardcodedNumber(formula: string): boolean {
  const body = stripStringLiterals(formula);
  return DIGIT.test(body.replace(WORD_TOKEN, ""));
}

export function classifyCell(formula: CellValue, value: CellValue): CellClass {
  if (!isFormula(formula)) {
    return value === null || value === "" ? "blank" : "input";
  }

  const body = stripStringLiterals(formula);
  if (hasWorkbookReference(body)) return "external";
  if (body.includes("!")) return "crossSheet";
  return hasHardcodedNumber(formula) ? "partial" : "formula";
}
