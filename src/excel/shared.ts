// Public building blocks the Excel adapter's sections share: formatting preset
// and number-format names, the selection summary shape, and address parsing.
// No Office.js state lives here; every export is a type or a pure function.

export type PresetName = "title" | "header" | "input" | "formula" | "result";
export type NumberFormatName = "whole" | "decimal" | "currency" | "percent";

export interface SelectionSummary {
  address: string;
  cells: number;
  formulas: number;
  errors: number;
  blanks: number;
}

// A change event carries every area it touched in one address, comma-separated
// and each area sheet-qualified: "Sheet1!A1:B2,Sheet1!D5". A quoted sheet name
// may hold a comma of its own, so the split walks the string instead of calling
// String.split, and an address with no comma comes back as its one area.
export function splitAreas(address: string): string[] {
  const areas: string[] = [];
  let start = 0;
  let quoted = false;
  for (let at = 0; at < address.length; at += 1) {
    const char = address[at];
    if (char === "'") quoted = !quoted;
    else if (char === "," && !quoted) {
      areas.push(address.slice(start, at));
      start = at + 1;
    }
  }
  areas.push(address.slice(start));
  return areas.filter((area) => area.trim() !== "");
}

// One area only: a multi-area address must go through splitAreas first, or
// everything before the last "!" is read as the sheet name. worksheet.getRange
// wants the local part, which is what the second half is.
export function parseAddress(address: string): {
  sheet: string;
  address: string;
} {
  const cut = address.lastIndexOf("!");
  if (cut < 0) return { sheet: "", address };
  return {
    sheet: address.slice(0, cut).replace(/^'|'$/g, "").replace(/''/g, "'"),
    address: address.slice(cut + 1),
  };
}
