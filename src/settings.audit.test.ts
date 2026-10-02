// The brand JSON the Brand tab imports: anything that is not a palette has to
// be refused whole. A half-applied file would leave a model wearing two brands.
import { describe, expect, it } from "vitest";
import { DEFAULT_SETTINGS, parsePalette, readStoredSettings } from "./settings";

describe("a brand file that is not one", () => {
  it("refuses valid JSON that is not an object", () => {
    expect(parsePalette("42")).toBeNull();
    expect(parsePalette("null")).toBeNull();
  });

  // A file, a localStorage value or a workbook setting carrying none of the
  // palette's own keys is not a partial palette, it is the wrong content: it
  // used to read as the shipped defaults, so the wrong .json imported as
  // "Palette imported" and a corrupted stored value silently replaced a brand.
  it("refuses an object carrying none of the palette's keys", () => {
    expect(parsePalette("{}")).toBeNull();
    expect(parsePalette('["#14213D"]')).toBeNull();
    expect(parsePalette('{"name":"plsfix","version":"2.6.2"}')).toBeNull();
  });

  it("still takes a palette that names only one of them", () => {
    // One known key is a partial palette: the rest fall back to the defaults.
    expect(parsePalette('{"accent":"#B27E54"}')).toEqual({
      ...DEFAULT_SETTINGS,
      accent: "#B27E54",
    });
    expect(parsePalette('{"autocolorOnEdit":true}')).toEqual({
      ...DEFAULT_SETTINGS,
      autocolorOnEdit: true,
    });
    // Unknown keys beside a known one are still ignored.
    expect(parsePalette('{"primary":"b27e54","unknown":1}')?.primary).toBe(
      "#B27E54",
    );
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
  // machine's stored copy is missing, unreadable or foreign. A refusal is a
  // fallback here and a sentence in the Brand tab's importer - never a throw.
  it("falls back to the shipped palette", () => {
    expect(readStoredSettings(null)).toEqual(DEFAULT_SETTINGS);
    expect(readStoredSettings("not json")).toEqual(DEFAULT_SETTINGS);
    expect(readStoredSettings('{"theme":"dark"}')).toEqual(DEFAULT_SETTINGS);
  });
});
