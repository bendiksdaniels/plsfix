import { describe, expect, it } from "vitest";
import {
  contrastText,
  currencyNumberFormat,
  DEFAULT_SETTINGS,
  deriveTheme,
  extractPaletteFromPixels,
  normalizeHex,
  parsePalette,
  serializeSettings,
  shade,
  tint,
} from "./settings";

describe("hex handling", () => {
  it("normalizes hex colors to #RRGGBB uppercase", () => {
    expect(normalizeHex("#b27e54")).toBe("#B27E54");
    expect(normalizeHex("b27e54")).toBe("#B27E54");
    expect(normalizeHex("#abc")).toBe("#AABBCC");
  });

  it("rejects invalid hex colors", () => {
    expect(normalizeHex("")).toBeNull();
    expect(normalizeHex("#12")).toBeNull();
    expect(normalizeHex("xyzxyz")).toBeNull();
  });
});

describe("color math", () => {
  it("tints toward white and shades toward black", () => {
    expect(tint("#000000", 0.5)).toBe("#808080");
    expect(shade("#FFFFFF", 0.5)).toBe("#808080");
  });

  it("picks readable text color for a background", () => {
    expect(contrastText("#282623")).toBe("#FFFFFF");
    expect(contrastText("#B27E54")).toBe("#FFFFFF");
    expect(contrastText("#E8EEF3")).toBe("#1A1918");
  });
});

describe("theme derivation", () => {
  it("derives all workbook colors from the palette", () => {
    const theme = deriveTheme(DEFAULT_SETTINGS);
    expect(theme.titleFill).toBe(DEFAULT_SETTINGS.primary);
    expect(theme.titleText).toBe(contrastText(DEFAULT_SETTINGS.primary));
    expect(theme.headerFill).toBe(tint(DEFAULT_SETTINGS.primary, 0.92));
    expect(theme.headerBorder).toBe(tint(DEFAULT_SETTINGS.primary, 0.55));
    expect(theme.resultFill).toBe(tint(DEFAULT_SETTINGS.accent, 0.86));
    expect(theme.resultBorder).toBe(shade(DEFAULT_SETTINGS.accent, 0.25));
    expect(theme.inputFont).toBe(DEFAULT_SETTINGS.input);
    expect(theme.formulaFont).toBe(DEFAULT_SETTINGS.formula);
    expect(theme.linkFont).toBe(DEFAULT_SETTINGS.link);
  });
});

describe("number formats", () => {
  it("builds a currency format around the symbol", () => {
    expect(currencyNumberFormat("€")).toBe("€ #,##0;[Red](€ #,##0);-");
    expect(currencyNumberFormat("$")).toBe("$ #,##0;[Red]($ #,##0);-");
  });

  it("drops the symbol when none is configured", () => {
    expect(currencyNumberFormat("")).toBe("#,##0;[Red](#,##0);-");
  });
});

describe("palette import and export", () => {
  it("round-trips settings through JSON", () => {
    const json = serializeSettings(DEFAULT_SETTINGS);
    expect(parsePalette(json)).toEqual(DEFAULT_SETTINGS);
  });

  it("merges partial palettes over defaults and normalizes hex", () => {
    const parsed = parsePalette('{"primary":"b27e54","unknown":1}');
    expect(parsed).toEqual({ ...DEFAULT_SETTINGS, primary: "#B27E54" });
  });

  it("rejects malformed JSON and invalid colors", () => {
    expect(parsePalette("not json")).toBeNull();
    expect(parsePalette('{"primary":"nope"}')).toBeNull();
    expect(parsePalette('{"currency":"€€€€"}')).toBeNull();
  });
});

describe("stored settings", () => {
  it("falls back to defaults for missing or corrupt storage", async () => {
    const { readStoredSettings } = await import("./settings");
    expect(readStoredSettings(null)).toEqual(DEFAULT_SETTINGS);
    expect(readStoredSettings("{broken")).toEqual(DEFAULT_SETTINGS);
    const custom = { ...DEFAULT_SETTINGS, accent: "#112233" };
    expect(readStoredSettings(serializeSettings(custom))).toEqual(custom);
  });
});

describe("logo color extraction", () => {
  const pixel = (r: number, g: number, b: number, a = 255) => [r, g, b, a];

  it("returns dominant colors, skipping white background and transparency", () => {
    const data = new Uint8ClampedArray(
      [
        ...Array.from({ length: 10 }, () => pixel(178, 126, 84)),
        ...Array.from({ length: 5 }, () => pixel(40, 38, 35)),
        ...Array.from({ length: 3 }, () => pixel(255, 255, 255)),
        ...Array.from({ length: 2 }, () => pixel(255, 0, 0, 40)),
      ].flat(),
    );

    expect(extractPaletteFromPixels(data, 4)).toEqual(["#B27E54", "#282623"]);
  });

  it("caps the number of returned colors", () => {
    const data = new Uint8ClampedArray(
      [
        ...Array.from({ length: 4 }, () => pixel(178, 126, 84)),
        ...Array.from({ length: 3 }, () => pixel(40, 38, 35)),
        ...Array.from({ length: 2 }, () => pixel(0, 87, 184)),
      ].flat(),
    );

    expect(extractPaletteFromPixels(data, 2)).toEqual(["#B27E54", "#282623"]);
  });
});
