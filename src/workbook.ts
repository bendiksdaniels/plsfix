// Two small workbook-wide reads: the contents-sheet rows for every visible
// sheet, and which defined names broke when Excel rewrote their formula to
// #REF!. Pure: no Office.js. Invariant: a hidden sheet is left off the
// contents sheet rather than linked and unreachable.
export interface TocRow {
  index: number;
  name: string;
  target: string;
}

// A sheet reference is always quoted, so spaces and punctuation need no special
// case; an apostrophe inside the name is doubled, as Excel writes it itself.
function sheetTarget(name: string): string {
  return `'${name.replace(/'/g, "''")}'!A1`;
}

// Hidden sheets are hidden on purpose: a contents sheet that links to them
// would both fail to jump and expose the workbook's back rooms.
export function tocRows(
  sheets: { name: string; visibility: string }[],
): TocRow[] {
  return sheets
    .filter((sheet) => sheet.visibility.toLowerCase() === "visible")
    .map((sheet, position) => ({
      index: position + 1,
      name: sheet.name,
      target: sheetTarget(sheet.name),
    }));
}

// A name survives the deletion of what it pointed at; Excel rewrites its
// reference to #REF! and every formula using the name breaks with it.
export function brokenNames(
  items: { name: string; formula: string }[],
): string[] {
  return items
    .filter((item) => item.formula.includes("#REF!"))
    .map((item) => item.name);
}
