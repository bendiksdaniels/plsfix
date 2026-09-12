// A1 references inside a formula string: where each token sits, what it points
// at, and how to write a changed one back in its place. Pure - no Office.js,
// no workbook. Owns the scanner only: whether a reference should move is the
// caller's rule (src/paste.ts duplicateFormula).
// Invariant: text inside double quotes and inside a structured reference's
// brackets is never a reference, and a token that a name or a "(" runs into is
// part of that name or that function call, not an address.

/** Excel's own grid, so a whole-column reference has a size to be tested. */
export const GRID_ROWS = 1_048_576;
export const GRID_COLUMNS = 16_384;

const MAX_COLUMN_LETTERS = 3;
const MAX_ROW_DIGITS = 7;
const LETTER_BASE = 26;
const UPPER_A = 65;
// What a defined name is made of: a token these run into is part of the name.
// Any script's letters count, so an accented sheet or range name reads as one
// word instead of breaking apart around its non-ASCII characters.
const NAME_CHAR = /[\p{L}\p{N}_.]/u;
const LETTER = /[A-Za-z]/;
const DIGIT = /[0-9]/;

/** One corner of a reference. A null row is "A:A", a null column is "1:1". */
export interface RefPart {
  column: number | null;
  columnAbsolute: boolean;
  row: number | null;
  rowAbsolute: boolean;
}

export interface FormulaRef {
  /** Index of the first character of the token, sheet prefix included. */
  start: number;
  /** Index one past the last character. */
  end: number;
  /** The token exactly as written. */
  text: string;
  /** The sheet prefix as written, "!" included; "" when there is none. */
  prefix: string;
  /** The prefix's sheet name, unquoted; "" when there is no prefix. */
  sheet: string;
  from: RefPart;
  /** The second corner, or `from` again when the token names one cell. */
  to: RefPart;
  pair: boolean;
}

function columnNumber(letters: string): number {
  let value = 0;
  for (const letter of letters.toUpperCase()) {
    value = value * LETTER_BASE + (letter.charCodeAt(0) - UPPER_A + 1);
  }
  return value;
}

function columnName(index: number): string {
  let name = "";
  let left = index;
  while (left >= 0) {
    name = String.fromCharCode(UPPER_A + (left % LETTER_BASE)) + name;
    left = Math.floor(left / LETTER_BASE) - 1;
  }
  return name;
}

/** Index just past the closing quote of the string starting at `at`. */
function skipString(text: string, at: number): number {
  let index = at + 1;
  while (index < text.length) {
    if (text[index] === '"') {
      // "" inside a string is one literal quote, not the end of it.
      if (text[index + 1] === '"') {
        index += 2;
        continue;
      }
      return index + 1;
    }
    index += 1;
  }
  return text.length;
}

/** Index just past the matching "]", or -1 when the group never closes. */
function skipBrackets(text: string, at: number): number {
  let depth = 0;
  for (let index = at; index < text.length; index += 1) {
    if (text[index] === "[") depth += 1;
    else if (text[index] === "]") {
      depth -= 1;
      if (depth === 0) return index + 1;
    }
  }
  return -1;
}

interface QuotedName {
  name: string;
  end: number;
}

function readQuotedName(text: string, at: number): QuotedName | null {
  let name = "";
  let index = at + 1;
  while (index < text.length) {
    if (text[index] === "'") {
      if (text[index + 1] === "'") {
        name += "'";
        index += 2;
        continue;
      }
      return { name, end: index + 1 };
    }
    name += text[index];
    index += 1;
  }
  return null;
}

interface PrefixRead {
  prefix: string;
  sheet: string;
  end: number;
}

// "Sheet2!", "'P&L 2025'!", "[Book1.xlsx]Sheet1!". The "!" is what makes a
// prefix: without it this is a function name or a defined name.
function readPrefix(text: string, at: number): PrefixRead | null {
  let index = at;
  let workbook = "";
  if (text[index] === "[") {
    const close = skipBrackets(text, index);
    if (close < 0) return null;
    workbook = text.slice(index, close);
    index = close;
  }

  let name = "";
  if (text[index] === "'") {
    const quoted = readQuotedName(text, index);
    if (!quoted) return null;
    name = quoted.name;
    index = quoted.end;
  } else {
    while (NAME_CHAR.test(text[index] ?? "")) {
      name += text[index];
      index += 1;
    }
  }

  if (text[index] !== "!" || (name === "" && workbook === "")) return null;
  return {
    prefix: text.slice(at, index + 1),
    sheet: workbook + name,
    end: index + 1,
  };
}

interface PartRead {
  part: RefPart;
  end: number;
}

// "$A$1", "A1", "$A" (half of "$A:$A") or "1" (half of "1:1"). Null for
// anything that is not an address at all, and for one past the grid: "XFE1"
// and "A1048577" are names as far as this scanner is concerned.
function readPart(text: string, at: number): PartRead | null {
  let index = at;
  let columnAbsolute = false;
  if (text[index] === "$") {
    columnAbsolute = true;
    index += 1;
  }

  let letters = "";
  while (
    letters.length < MAX_COLUMN_LETTERS &&
    LETTER.test(text[index] ?? "")
  ) {
    letters += text[index];
    index += 1;
  }
  // A fourth letter means a name ("Table"), never a column.
  if (LETTER.test(text[index] ?? "")) return null;

  let rowAbsolute = false;
  if (letters === "") {
    // The lone "$" belonged to the row all along: "$1" of "$1:$1".
    rowAbsolute = columnAbsolute;
    columnAbsolute = false;
  } else if (text[index] === "$") {
    rowAbsolute = true;
    index += 1;
  }

  let digits = "";
  while (DIGIT.test(text[index] ?? "")) {
    digits += text[index];
    index += 1;
  }
  if (letters === "" && digits === "") return null;
  if (digits.length > MAX_ROW_DIGITS) return null;

  const column = letters === "" ? null : columnNumber(letters) - 1;
  const row = digits === "" ? null : Number(digits) - 1;
  if (column !== null && column >= GRID_COLUMNS) return null;
  if (row !== null && (row < 0 || row >= GRID_ROWS)) return null;
  return { part: { column, columnAbsolute, row, rowAbsolute }, end: index };
}

// Two corners have to name the same kind of thing: two cells, two whole
// columns or two whole rows. "A1:B" is not an address we understand.
function sameKind(first: RefPart, second: RefPart): boolean {
  return (
    (first.column === null) === (second.column === null) &&
    (first.row === null) === (second.row === null)
  );
}

function readReference(text: string, at: number): FormulaRef | null {
  const prefix = readPrefix(text, at);
  const first = readPart(text, prefix ? prefix.end : at);
  if (!first) return null;

  let last = first;
  let pair = false;
  if (text[first.end] === ":") {
    const second = readPart(text, first.end + 1);
    if (second && sameKind(first.part, second.part)) {
      last = second;
      pair = true;
    }
  }
  // A bare "A" or "3" only names something as half of a pair.
  if (!pair && (first.part.column === null || first.part.row === null)) {
    return null;
  }

  const before = text[at - 1] ?? "";
  const after = text[last.end] ?? "";
  if (NAME_CHAR.test(before) || before === "$") return null;
  if (NAME_CHAR.test(after) || after === "(") return null;

  return {
    start: at,
    end: last.end,
    text: text.slice(at, last.end),
    prefix: prefix?.prefix ?? "",
    sheet: prefix?.sheet ?? "",
    from: first.part,
    to: last.part,
    pair,
  };
}

/** Every A1 reference in a formula, in the order they are written. */
export function findReferences(formula: string): FormulaRef[] {
  const refs: FormulaRef[] = [];
  let index = 0;
  while (index < formula.length) {
    if (formula[index] === '"') {
      index = skipString(formula, index);
      continue;
    }
    const ref = readReference(formula, index);
    if (ref) {
      refs.push(ref);
      index = ref.end;
      continue;
    }
    // Not the start of a reference, so a "[" here opens a structured
    // reference's body: everything inside it is a column name, not an address.
    if (formula[index] === "[") {
      const close = skipBrackets(formula, index);
      index = close < 0 ? index + 1 : close;
      continue;
    }
    index += 1;
  }
  return refs;
}

export function formatPart(part: RefPart): string {
  const column =
    part.column === null
      ? ""
      : (part.columnAbsolute ? "$" : "") + columnName(part.column);
  const row =
    part.row === null
      ? ""
      : (part.rowAbsolute ? "$" : "") + String(part.row + 1);
  return column + row;
}

/** The same token with new corners: sheet prefix and ":" kept as written. */
export function formatReference(
  ref: FormulaRef,
  from: RefPart,
  to: RefPart,
): string {
  const head = ref.prefix + formatPart(from);
  return ref.pair ? `${head}:${formatPart(to)}` : head;
}

/**
 * Rewrites the references a caller returns a replacement for and leaves every
 * other character of the formula exactly where it was.
 */
export function rewriteReferences(
  formula: string,
  rewrite: (ref: FormulaRef) => string | null,
): string {
  let out = "";
  let at = 0;
  for (const ref of findReferences(formula)) {
    const replacement = rewrite(ref);
    if (replacement === null) continue;
    out += formula.slice(at, ref.start) + replacement;
    at = ref.end;
  }
  return out + formula.slice(at);
}
