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

function stepSection(section: string, delta: 1 | -1): string {
  const literal = scanLiterals(section);

  let start = -1;
  for (let index = 0; index < section.length; index += 1) {
    if (!literal[index] && PLACEHOLDERS.includes(section[index] ?? "")) {
      start = index;
      break;
    }
  }
  // No digit placeholders: "General", "@" and date codes step nowhere.
  if (start < 0) return section;

  let end = start;
  while (
    end + 1 < section.length &&
    !literal[end + 1] &&
    NUMERIC_BODY.includes(section[end + 1] ?? "")
  ) {
    end += 1;
  }
  // A trailing comma scales by a thousand; it is not part of the decimals.
  while (end > start && section[end] === ",") end -= 1;

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
  if (delta === 1) return `${section.slice(0, end + 1)}0${section.slice(end + 1)}`;
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

export function buildCagrFormula(
  firstRef: string,
  lastRef: string,
  periods: number,
): string {
  return `=(${lastRef}/${firstRef})^(1/${periods})-1`;
}
