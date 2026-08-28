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
