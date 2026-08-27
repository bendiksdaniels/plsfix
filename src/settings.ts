export interface BrandSettings {
  primary: string;
  accent: string;
  input: string;
  formula: string;
  link: string;
  external: string;
  partial: string;
  font: string;
  currency: string;
  autocolorOnEdit: boolean;
}

export interface WorkbookTheme {
  titleFill: string;
  titleText: string;
  headerFill: string;
  headerBorder: string;
  resultFill: string;
  resultBorder: string;
  inputFont: string;
  formulaFont: string;
  linkFont: string;
  externalFont: string;
  partialFont: string;
}

export const DEFAULT_SETTINGS: BrandSettings = {
  primary: "#282623",
  accent: "#B27E54",
  input: "#0057B8",
  formula: "#1F1D1B",
  link: "#17823B",
  external: "#C00000",
  partial: "#7A3E9D",
  font: "Aptos",
  currency: "€",
  autocolorOnEdit: false,
};

const COLOR_KEYS = [
  "primary",
  "accent",
  "input",
  "formula",
  "link",
  "external",
  "partial",
] as const;

export function normalizeHex(value: string): string | null {
  let hex = value.trim().replace(/^#/, "");
  if (/^[0-9a-f]{3}$/i.test(hex)) {
    hex = hex.replace(/./g, (char) => char + char);
  }
  if (!/^[0-9a-f]{6}$/i.test(hex)) return null;
  return `#${hex.toUpperCase()}`;
}

function channels(hex: string): [number, number, number] {
  return [
    parseInt(hex.slice(1, 3), 16),
    parseInt(hex.slice(3, 5), 16),
    parseInt(hex.slice(5, 7), 16),
  ];
}

function toHex(r: number, g: number, b: number): string {
  const part = (channel: number) =>
    Math.round(channel).toString(16).padStart(2, "0");
  return `#${part(r)}${part(g)}${part(b)}`.toUpperCase();
}

function mix(hex: string, target: number, ratio: number): string {
  const [r, g, b] = channels(hex);
  return toHex(
    r + (target - r) * ratio,
    g + (target - g) * ratio,
    b + (target - b) * ratio,
  );
}

export function tint(hex: string, ratio: number): string {
  return mix(hex, 255, ratio);
}

export function shade(hex: string, ratio: number): string {
  return mix(hex, 0, ratio);
}

export function contrastText(background: string): string {
  const [r, g, b] = channels(background);
  const brightness = (299 * r + 587 * g + 114 * b) / 1000;
  return brightness >= 150 ? "#1A1918" : "#FFFFFF";
}

export function deriveTheme(settings: BrandSettings): WorkbookTheme {
  return {
    titleFill: settings.primary,
    titleText: contrastText(settings.primary),
    headerFill: tint(settings.primary, 0.92),
    headerBorder: tint(settings.primary, 0.55),
    resultFill: tint(settings.accent, 0.86),
    resultBorder: shade(settings.accent, 0.25),
    inputFont: settings.input,
    formulaFont: settings.formula,
    linkFont: settings.link,
    externalFont: settings.external,
    partialFont: settings.partial,
  };
}

export function currencyNumberFormat(symbol: string): string {
  if (!symbol) return "#,##0;[Red](#,##0);-";
  return `${symbol} #,##0;[Red](${symbol} #,##0);-`;
}

export function serializeSettings(settings: BrandSettings): string {
  return JSON.stringify(settings, null, 2);
}

export function parsePalette(json: string): BrandSettings | null {
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch {
    return null;
  }
  if (typeof raw !== "object" || raw === null) return null;

  const source = raw as Record<string, unknown>;
  const settings: BrandSettings = { ...DEFAULT_SETTINGS };

  for (const key of COLOR_KEYS) {
    const value = source[key];
    if (value === undefined) continue;
    if (typeof value !== "string") return null;
    const hex = normalizeHex(value);
    if (!hex) return null;
    settings[key] = hex;
  }

  if (source.font !== undefined) {
    if (typeof source.font !== "string" || !source.font.trim()) return null;
    settings.font = source.font.trim();
  }

  if (source.currency !== undefined) {
    if (typeof source.currency !== "string" || source.currency.length > 3) {
      return null;
    }
    settings.currency = source.currency;
  }

  if (source.autocolorOnEdit !== undefined) {
    if (typeof source.autocolorOnEdit !== "boolean") return null;
    settings.autocolorOnEdit = source.autocolorOnEdit;
  }

  return settings;
}

export function readStoredSettings(raw: string | null): BrandSettings {
  if (!raw) return { ...DEFAULT_SETTINGS };
  return parsePalette(raw) ?? { ...DEFAULT_SETTINGS };
}

let active: BrandSettings = { ...DEFAULT_SETTINGS };
let activeThemeCache: WorkbookTheme = deriveTheme(active);

export function getActiveSettings(): BrandSettings {
  return active;
}

export function activeTheme(): WorkbookTheme {
  return activeThemeCache;
}

export function setActiveSettings(settings: BrandSettings): void {
  active = { ...settings };
  activeThemeCache = deriveTheme(active);
}

// Logos usually sit on white, so near-white pixels are treated as background.
export function extractPaletteFromPixels(
  data: Uint8ClampedArray,
  maxColors: number,
): string[] {
  interface Bucket {
    count: number;
    r: number;
    g: number;
    b: number;
  }
  const buckets = new Map<number, Bucket>();

  for (let i = 0; i + 3 < data.length; i += 4) {
    const r = data[i]!;
    const g = data[i + 1]!;
    const b = data[i + 2]!;
    const a = data[i + 3]!;
    if (a < 128) continue;
    if (r >= 246 && g >= 246 && b >= 246) continue;

    const key = ((r >> 4) << 8) | ((g >> 4) << 4) | (b >> 4);
    const bucket = buckets.get(key) ?? { count: 0, r: 0, g: 0, b: 0 };
    bucket.count += 1;
    bucket.r += r;
    bucket.g += g;
    bucket.b += b;
    buckets.set(key, bucket);
  }

  return [...buckets.values()]
    .sort((a, b) => b.count - a.count)
    .slice(0, maxColors)
    .map((bucket) =>
      toHex(bucket.r / bucket.count, bucket.g / bucket.count, bucket.b / bucket.count),
    );
}
