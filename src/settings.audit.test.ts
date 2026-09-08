// The brand JSON the Brand tab imports: anything that is not a palette has to
// be refused whole. A half-applied file would leave a model wearing two brands.
import { describe, expect, it } from "vitest";
import { DEFAULT_SETTINGS, parsePalette, readStoredSettings } from "./settings";

describe("a brand file that is not one", () => {
  it("refuses valid JSON that is not an object", () => {
    expect(parsePalette("42")).toBeNull();
    expect(parsePalette("null")).toBeNull();
  });

  // Pinned by src/settings.test.ts: an object with no palette key in it reads
  // as the shipped palette rather than as a refusal. That is also what the
  // Brand tab's importer does with the wrong .json - it reports "Palette
  // imported" and the brand becomes the default one.
  it("reads any other object as the shipped palette", () => {
    expect(parsePalette("{}")).toEqual(DEFAULT_SETTINGS);
    expect(parsePalette('["#14213D"]')).toEqual(DEFAULT_SETTINGS);
    expect(parsePalette('{"name":"plsfix","version":"2.6.2"}')).toEqual(
      DEFAULT_SETTINGS,
    );
    // One known key is a partial palette: the rest fall back to the defaults.
    expect(parsePalette('{"accent":"#B27E54"}')).toEqual({
      ...DEFAULT_SETTINGS,
      accent: "#B27E54",
    });
  });

  it("refuses a colour that is not a string", () => {
    expect(parsePalette('{"accent":16711680}')).toBeNull();
    expect(parsePalette('{"input":null}')).toBeNull();
  });

  it("refuses a font that is missing or blank", () => {
    expect(parsePalette('{"font":12}')).toBeNull();
    expect(parsePalette('{"font":"   "}')).toBeNull();
    expect(parsePalette('{"font":" Georgia "}')?.font).toBe("Georgia");
  });

  it("refuses a language the pane has no number style for", () => {
    expect(parsePalette('{"language":"de"}')).toBeNull();
  });

  // The pane boots on the shipped palette rather than on nothing when the
  // machine's stored copy is missing or unreadable.
  it("falls back to the shipped palette", () => {
    expect(readStoredSettings(null)).toEqual(DEFAULT_SETTINGS);
    expect(readStoredSettings("not json")).toEqual(DEFAULT_SETTINGS);
  });
});
