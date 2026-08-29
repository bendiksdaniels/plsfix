import {
  type BrandSettings,
  contrastText,
  currencyNumberFormat,
  deriveTheme,
} from "./settings";

export type NumberCycleFamily =
  "general" | "currency" | "percent" | "multiple" | "date";
export type RowStyleKind = "title" | "result" | "item";

export type NumberCycles = Record<NumberCycleFamily, string[]>;
export type RowStyleCycles = Record<RowStyleKind, StyleSpec[]>;

export interface BorderSpec {
  style: "double" | "continuous";
  color: string;
}

export interface StyleSpec {
  fill?: string;
  fontColor?: string;
  bold?: boolean;
  topBorder?: BorderSpec | null;
  bottomBorder?: BorderSpec | null;
}

export interface CellStyle {
  fill: string | null;
  fontColor: string;
  bold: boolean;
}

// Sentinel fill: Excel has no "no fill" color, the caller clears the range instead.
export const CLEAR_FILL = "clear";

function financial(body: string): string {
  return `${body};[Red](${body});-`;
}

function withSymbol(symbol: string, digits: string): string {
  return symbol ? `${symbol} ${digits}` : digits;
}

export function buildNumberCycles(settings: BrandSettings): NumberCycles {
  const { currency } = settings;

  return {
    general: [financial("#,##0"), financial("#,##0.0"), financial("#,##0.00")],
    // Trailing comma divides the displayed value by a thousand.
    currency: [
      currencyNumberFormat(currency),
      financial(withSymbol(currency, "#,##0.0")),
      financial(withSymbol(currency, "#,##0,")),
    ],
    percent: [financial("0.0%"), financial("0%"), financial("0.00%")],
    // "x" is quoted so Excel keeps it as a literal rather than a format code.
    multiple: [financial('0.0"x"'), financial('0.00"x"')],
    date: ["dd.mm.yyyy", "mmm-yy", "yyyy"],
  };
}

// Excel rewrites plain currency symbols into locale-tagged codes on read-back
// (e.g. "€ #,##0" comes back as "[$€-x-euro2] #,##0"), so cycle matching must
// compare canonical forms while still writing the clean literal format.
export function canonicalNumberFormat(format: string): string {
  return format.replace(/\[\$([^-\]]+)(-[^\]]*)?\]/g, "$1");
}

export function nextInCycle(current: string, cycle: string[]): string {
  const canonical = canonicalNumberFormat(current);
  const index = cycle.findIndex(
    (entry) => canonicalNumberFormat(entry) === canonical,
  );
  // -1 for formats we did not apply, which steps to entry 0.
  const next = cycle[(index + 1) % cycle.length];
  return next ?? current;
}

export function buildRowStyleCycles(settings: BrandSettings): RowStyleCycles {
  const theme = deriveTheme(settings);
  const rule: BorderSpec = { style: "continuous", color: theme.headerBorder };
  const total: BorderSpec = { style: "double", color: theme.resultBorder };

  return {
    title: [
      {
        fill: theme.titleFill,
        fontColor: theme.titleText,
        bold: true,
        topBorder: null,
        bottomBorder: null,
      },
      {
        fill: theme.headerFill,
        fontColor: theme.formulaFont,
        bold: true,
        topBorder: null,
        bottomBorder: rule,
      },
      {
        fill: CLEAR_FILL,
        fontColor: settings.accent,
        bold: true,
        topBorder: null,
        bottomBorder: { style: "continuous", color: settings.accent },
      },
    ],
    result: [
      {
        fill: theme.resultFill,
        fontColor: theme.formulaFont,
        bold: true,
        topBorder: total,
        bottomBorder: null,
      },
      {
        fill: CLEAR_FILL,
        fontColor: theme.formulaFont,
        bold: true,
        topBorder: { style: "continuous", color: theme.resultBorder },
        bottomBorder: total,
      },
      {
        fill: settings.accent,
        fontColor: contrastText(settings.accent),
        bold: true,
        topBorder: null,
        bottomBorder: null,
      },
    ],
    item: [
      {
        fill: CLEAR_FILL,
        fontColor: theme.formulaFont,
        bold: false,
        topBorder: null,
        bottomBorder: null,
      },
      {
        fill: theme.headerFill,
        fontColor: theme.formulaFont,
        bold: false,
        topBorder: null,
        bottomBorder: null,
      },
      {
        fill: theme.resultFill,
        fontColor: theme.formulaFont,
        bold: false,
        topBorder: null,
        bottomBorder: null,
      },
    ],
  };
}

export function matchStyleIndex(
  current: CellStyle,
  variants: StyleSpec[],
): number {
  return variants.findIndex((variant) => {
    const fill = variant.fill === CLEAR_FILL ? null : variant.fill;
    if (variant.fill !== undefined && fill !== current.fill) return false;
    if (
      variant.fontColor !== undefined &&
      variant.fontColor !== current.fontColor
    ) {
      return false;
    }
    return variant.bold === undefined || variant.bold === current.bold;
  });
}

export function buildFillCycle(settings: BrandSettings): string[] {
  const theme = deriveTheme(settings);
  return [
    theme.headerFill,
    theme.resultFill,
    settings.accent,
    settings.primary,
    CLEAR_FILL,
  ];
}

export function buildFontCycle(settings: BrandSettings): string[] {
  const theme = deriveTheme(settings);
  return [
    theme.formulaFont,
    theme.inputFont,
    theme.linkFont,
    settings.accent,
    settings.primary,
  ];
}

// ---------------------------------------------------------------------------
// Border cycle: the looks a modeller draws on a block of a schedule, applied to
// the edges of the SELECTION (and its interior lines) rather than to each cell.
// ---------------------------------------------------------------------------

export const BORDER_EDGE_NAMES = [
  "top",
  "bottom",
  "left",
  "right",
  "insideHorizontal",
  "insideVertical",
] as const;

export type BorderEdgeName = (typeof BORDER_EDGE_NAMES)[number];

// One line of a border state: which edge of the selection it draws, and how.
export interface BorderLine extends BorderSpec {
  edge: BorderEdgeName;
  weight: "thin" | "medium";
}

export type BorderCycleState = BorderLine[];

// What the host reports for one edge. Excel spells these "Continuous", "Thin"
// and "#282623", and hands back an empty style for a range whose cells disagree.
export interface BorderReadout {
  style: string;
  weight: string;
  color: string;
}

// Edges the host cannot report are left out: a single cell has no inside lines.
export type BorderReadouts = Partial<Record<BorderEdgeName, BorderReadout>>;

const NO_LINE = "none";
const OUTLINE_EDGES: BorderEdgeName[] = ["top", "bottom", "left", "right"];

// Six looks, in the order a schedule grows: nothing, the rule under a row, the
// heavier total rule, the double result underline, a box around the block and
// the full grid. Every line is drawn in the brand primary.
export function buildBorderCycle(settings: BrandSettings): BorderCycleState[] {
  const color = settings.primary;
  const thin = (edge: BorderEdgeName): BorderLine => ({
    edge,
    style: "continuous",
    weight: "thin",
    color,
  });
  const outline = OUTLINE_EDGES.map(thin);

  return [
    [],
    [thin("bottom")],
    [{ edge: "bottom", style: "continuous", weight: "medium", color }],
    [thin("top"), { edge: "bottom", style: "double", weight: "thin", color }],
    outline,
    [...outline, thin("insideHorizontal"), thin("insideVertical")],
  ];
}

// Excel reports style, weight and colour in its own spelling ("Continuous",
// "Thin"), so matching compares canonical forms the way number formats do.
function canonicalBorder(value: string): string {
  return value.trim().toLowerCase();
}

function drawnAs(line: BorderLine, read: BorderReadout): boolean {
  if (canonicalBorder(read.style) !== line.style) return false;
  if (canonicalBorder(read.color) !== canonicalBorder(line.color)) return false;
  // A double rule is drawn at Excel's own weight, so only single rules match it.
  return (
    line.style === "double" || canonicalBorder(read.weight) === line.weight
  );
}

function isState(state: BorderCycleState, current: BorderReadouts): boolean {
  for (const edge of BORDER_EDGE_NAMES) {
    const read = current[edge];
    // An edge the host cannot report is no evidence against the state.
    if (!read) continue;
    const line = state.find((entry) => entry.edge === edge);
    if (line ? !drawnAs(line, read) : canonicalBorder(read.style) !== NO_LINE) {
      return false;
    }
  }
  return true;
}

// The state the range is already in. Anything we did not draw - a hand-drawn
// line, a range whose cells disagree - reads as state 1 (no borders), so the
// next press starts our cycle instead of wiping a look we do not own. Matching
// runs from the back because on a single cell a box and a grid are the same
// drawing: the more specific state wins, and the cycle still comes home.
export function matchBorderIndex(
  current: BorderReadouts,
  states: BorderCycleState[],
): number {
  for (let index = states.length - 1; index >= 0; index -= 1) {
    const state = states[index];
    if (state && isState(state, current)) return index;
  }
  return 0;
}

// ---------------------------------------------------------------------------
// Size cycles: the row heights and column widths a model uses, stepped one
// press at a time. Both lists are in POINTS, the unit Office.js takes for
// `range.format.rowHeight` and `range.format.columnWidth` - not Excel's
// character-count column width, whose 8.43-character default is the 64 points
// this cycle starts on. The numbers are the ones Excel's own Row Height and
// Column Width dialogs show.
// ---------------------------------------------------------------------------

export interface SizeCycles {
  rowHeight: number[];
  columnWidth: number[];
}

// Rows start on Excel's default 15 pt and climb through the bands a model
// uses - a roomier line, a header, a title - before coming home. Columns widen
// for labels and end narrow for a spacer column; Excel's own default column is
// 8.43 characters, which is 48 pt, so the last rung is the one a default column
// steps home to rather than the first.
export function buildSizeCycles(): SizeCycles {
  return {
    rowHeight: [15, 18, 21, 24, 30],
    columnWidth: [64, 80, 96, 120, 48],
  };
}

// Excel stores sizes in points but snaps them to whole screen pixels, so a
// height we wrote as 21 can read back as 20.75: a size counts as one of ours
// when it is within half a point of it.
export function nextSize(
  current: number,
  cycle: number[],
  tolerance = 0.5,
): number {
  const index = cycle.findIndex(
    (entry) => Math.abs(entry - current) <= tolerance,
  );
  // -1 for a size we did not set - a hand-dragged row - which steps to entry 0.
  const next = cycle[(index + 1) % cycle.length];
  return next ?? current;
}
