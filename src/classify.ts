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
  if (body.includes("[")) return "external";
  if (body.includes("!")) return "crossSheet";
  return hasHardcodedNumber(formula) ? "partial" : "formula";
}
