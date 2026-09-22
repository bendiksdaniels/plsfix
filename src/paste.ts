// The pure maths behind the fill and paste actions: how far a fast fill
// reaches, the formula strings a paste writes (sign flip, IFERROR guard, CAGR,
// rounding), how a number format's decimals step, and how a source grid tiles
// over a larger target. Reference rewriting is src/formula-duplicate.ts.
// Invariant: every function here is a string or grid transform - what a cell
// currently holds goes in, what it should hold comes out, nothing is read back.

import { type CellValue, isFormula } from "./model";

function isBlank(value: CellValue): boolean {
  return value === null || value === undefined || value === "";
}

// Each inner array is one neighbour line ordered away from the fill origin, so
// the fill reaches as far as the longest unbroken run of data beside it.
export function detectFillExtent(neighbors: CellValue[][]): number {
  let longest = 0;
  for (const line of neighbors) {
    let run = 0;
    while (run < line.length && !isBlank(line[run] ?? null)) run += 1;
    if (run > longest) longest = run;
  }
  return longest;
}

// Walks from an opening parenthesis to its match, skipping quoted text.
// Returns -1 when it never closes.
function matchParen(body: string, open: number): number {
  let depth = 0;
  let quoted = false;

  for (let index = open; index < body.length; index += 1) {
    const char = body[index];
    if (quoted) {
      if (char === '"') quoted = false;
      continue;
    }
    if (char === '"') quoted = true;
    else if (char === "(") depth += 1;
    else if (char === ")") {
      depth -= 1;
      if (depth === 0) return index;
    }
  }

  return -1;
}

export function flipSign(cells: CellValue[][]): CellValue[][] {
  return cells.map((row) =>
    row.map((cell) => {
      if (typeof cell === "number") return -cell;
      if (!isFormula(cell)) return cell;

      const body = cell.slice(1);
      // Only our own wrap unwraps: "=-(A1)+B1" closes early and gets wrapped.
      if (body.startsWith("-(") && matchParen(body, 1) === body.length - 1) {
        return `=${body.slice(2, -1)}`;
      }
      return `=-(${body})`;
    }),
  );
}

const PLACEHOLDERS = "0#?";
const NUMERIC_BODY = "0#?,.";

// Quoted text, bracket codes like [Red] or [$€-x-euro2] and backslash escapes
// are literals: a ";" or "." inside them is not format syntax.
function scanLiterals(section: string): boolean[] {
  const literal = new Array<boolean>(section.length).fill(false);
  let quoted = false;
  let bracketed = false;

  for (let index = 0; index < section.length; index += 1) {
    const char = section[index];
    if (quoted) {
      literal[index] = true;
      if (char === '"') quoted = false;
    } else if (bracketed) {
      literal[index] = true;
      if (char === "]") bracketed = false;
    } else if (char === "\\") {
      literal[index] = true;
      if (index + 1 < section.length) {
        literal[index + 1] = true;
        index += 1;
      }
    } else if (char === '"') {
      literal[index] = true;
      quoted = true;
    } else if (char === "[") {
      literal[index] = true;
      bracketed = true;
    }
  }

  return literal;
}

function splitSections(format: string): string[] {
  const literal = scanLiterals(format);
  const sections: string[] = [];
  let current = "";

  for (let index = 0; index < format.length; index += 1) {
    if (format[index] === ";" && !literal[index]) {
      sections.push(current);
      current = "";
      continue;
    }
    current += format[index];
  }
  sections.push(current);

  return sections;
}

interface NumericBody {
  start: number;
  end: number; // last character before the trailing commas
  thousands: number; // trailing commas, each dividing what is printed by 1,000
}

// The run of digit placeholders inside one section - "#,##0.00" within
// "€ #,##0.00_);[Red](#,##0.00)" - or null when the section prints no digits at
// all: "General", "@" and date codes have nothing to step or count.
function numericBody(section: string, literal: boolean[]): NumericBody | null {
  let start = -1;
  for (let index = 0; index < section.length; index += 1) {
    if (!literal[index] && PLACEHOLDERS.includes(section[index] ?? "")) {
      start = index;
      break;
    }
  }
  if (start < 0) return null;

  let end = start;
  while (
    end + 1 < section.length &&
    !literal[end + 1] &&
    NUMERIC_BODY.includes(section[end + 1] ?? "")
  ) {
    end += 1;
  }
  // A trailing comma scales by a thousand; it is not part of the decimals.
  let thousands = 0;
  while (end > start && section[end] === ",") {
    thousands += 1;
    end -= 1;
  }
  return { start, end, thousands };
}

function stepSection(section: string, delta: 1 | -1): string {
  const literal = scanLiterals(section);
  const body = numericBody(section, literal);
  if (!body) return section;
  const { start, end } = body;

  let dot = -1;
  for (let index = start; index <= end; index += 1) {
    if (section[index] === ".") {
      dot = index;
      break;
    }
  }

  if (dot < 0) {
    if (delta === -1) return section;
    return `${section.slice(0, end + 1)}.0${section.slice(end + 1)}`;
  }
  if (delta === 1)
    return `${section.slice(0, end + 1)}0${section.slice(end + 1)}`;
  // Dropping the last decimal drops the separator with it.
  if (end - dot <= 1) return section.slice(0, dot) + section.slice(end + 1);
  return section.slice(0, end) + section.slice(end + 1);
}

export function stepDecimals(format: string, delta: 1 | -1): string {
  return splitSections(format)
    .map((section) => stepSection(section, delta))
    .join(";");
}

// Returns the guarded expression when the whole formula is one IFERROR call,
// null otherwise. A comma inside a nested call or a string is not the separator.
function unwrapIfError(body: string): string | null {
  const head = /^IFERROR\s*\(/i.exec(body);
  if (!head) return null;

  const open = head[0].length - 1;
  if (matchParen(body, open) !== body.length - 1) return null;

  let depth = 0;
  let quoted = false;
  for (let index = open; index < body.length - 1; index += 1) {
    const char = body[index];
    if (quoted) {
      if (char === '"') quoted = false;
    } else if (char === '"') quoted = true;
    else if (char === "(") depth += 1;
    else if (char === ")") depth -= 1;
    else if (char === "," && depth === 1) return body.slice(open + 1, index);
  }

  return null;
}

export function toggleIfError(
  cells: CellValue[][],
  fallback: string,
): CellValue[][] {
  return cells.map((row) =>
    row.map((cell) => {
      if (!isFormula(cell)) return cell;
      const body = cell.slice(1);
      const guarded = unwrapIfError(body);
      return guarded === null ? `=IFERROR(${body},${fallback})` : `=${guarded}`;
    }),
  );
}

// The same read toggleIfError makes, without building the output grid: what
// the receipt (src/excel/formulas.ts) counts before it writes.
export function countIfErrorToggle(cells: CellValue[][]): {
  added: number;
  stripped: number;
} {
  let added = 0;
  let stripped = 0;
  for (const row of cells) {
    for (const cell of row) {
      if (!isFormula(cell)) continue;
      if (unwrapIfError(cell.slice(1)) === null) added += 1;
      else stripped += 1;
    }
  }
  return { added, stripped };
}

export function buildCagrFormula(
  firstRef: string,
  lastRef: string,
  periods: number,
): string {
  return `=(${lastRef}/${firstRef})^(1/${periods})-1`;
}

// The precision a number format prints, read back as a rounding precision on
// the stored value: "#,##0.00" is 2, a percentage adds the two places Excel
// shifts before printing ("0.0%" is 3), and each thousands comma takes three
// away ("#,##0,," rounds to millions). Null when the format prints no digits,
// which leaves the caller's own default standing. Only the first section is
// read: that is the one a positive number prints through.
export function formatDecimals(format: string): number | null {
  const section = splitSections(format)[0] ?? "";
  const literal = scanLiterals(section);
  const body = numericBody(section, literal);
  if (!body) return null;

  const digits = section.slice(body.start, body.end + 1);
  const dot = digits.indexOf(".");
  const decimals =
    dot < 0
      ? 0
      : [...digits.slice(dot + 1)].filter((char) => PLACEHOLDERS.includes(char))
          .length;
  const percent = [...section].some(
    (char, index) => char === "%" && !literal[index],
  );
  return decimals + (percent ? 2 : 0) - 3 * body.thousands;
}

// "A1:A5" as "$A$1:$A$5", any sheet prefix left as it is. Every cell of a
// rounding group points at the whole group, so the reference has to survive
// being filled or copied down the column.
export function absoluteRef(address: string): string {
  const cut = address.lastIndexOf("!");
  const local = address.slice(cut + 1);
  return (
    address.slice(0, cut + 1) +
    local.replaceAll(/([A-Za-z]+)(\d+)/g, "$$$1$$$2")
  );
}

// =PLSFIX.ROUND($A$1:$A$5,2,0): the whole group as the first argument, so Excel
// recalculates every sibling cell whenever any value in it changes, plus the
// literal position this cell reads out of the allocation.
export function buildRoundFormula(
  rangeRef: string,
  index: number,
  decimals: number,
): string {
  return `=PLSFIX.ROUND(${rangeRef},${index},${decimals})`;
}

/**
 * A source grid repeated over a target of the given shape, the way Excel's own
 * copy tiles a smaller source across a larger destination: destination cell
 * (r, c) reads source (r mod rows, c mod columns), so the last tile is cut
 * short when the target is not a whole multiple of the source.
 */
export function tileGrid<T>(grid: T[][], rows: number, columns: number): T[][] {
  const sourceRows = grid.length;
  const sourceColumns = grid[0]?.length ?? 0;
  if (sourceRows === 0 || sourceColumns === 0) return [];
  return Array.from({ length: rows }, (_unused, row) =>
    Array.from(
      { length: columns },
      (_cell, column) => grid[row % sourceRows]![column % sourceColumns]!,
    ),
  );
}
