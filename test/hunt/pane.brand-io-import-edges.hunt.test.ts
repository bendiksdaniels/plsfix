// @vitest-environment jsdom
// readPaletteFile/isPaletteFile (src/pane/brand-io.ts) had no DIRECT unit
// test: only the DOM-driven brand-io.audit.test.ts exercises them, through
// the file picker. The underlying parsePalette (src/settings.ts) is well
// covered on its own, but this file's OWN gate - "does this JSON name even
// one of the palette's own keys at all" - had never been unit-tested in
// isolation. Odd values named in the brief: garbage, truncated, missing
// keys, a future version (extra keys this build has never heard of), a
// 2 MB file.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { DEFAULT_SETTINGS } from "../../src/settings";
import type { readPaletteFile as ReadPaletteFile } from "../../src/pane/brand-io";

let readPaletteFile: typeof ReadPaletteFile;

beforeEach(async () => {
  vi.resetModules();
  // ./shared's top-level createToast/installTabs need #toast/#tab-bar to
  // already exist - the real markup, the same way every other pane test
  // loads it, since brand-io.ts imports ./shared for its toast alone.
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  ({ readPaletteFile } = await import("../../src/pane/brand-io"));
});

describe("readPaletteFile: garbage and truncated input", () => {
  it("refuses text that is not JSON at all", () => {
    expect(readPaletteFile("this is not json")).toBeNull();
    expect(readPaletteFile("")).toBeNull();
  });

  it("refuses JSON truncated mid-object, not just non-JSON text", () => {
    const whole = JSON.stringify(DEFAULT_SETTINGS);
    // Cut it off partway through a value: still garbage to JSON.parse, but a
    // DIFFERENT failure shape than a plain non-JSON string.
    const truncated = whole.slice(0, Math.floor(whole.length / 2));
    expect(readPaletteFile(truncated)).toBeNull();
  });

  it("refuses a bare JSON scalar and a bare array", () => {
    expect(readPaletteFile("null")).toBeNull();
    expect(readPaletteFile("42")).toBeNull();
    expect(readPaletteFile('"#123456"')).toBeNull();
    expect(readPaletteFile("[]")).toBeNull();
    expect(readPaletteFile('["primary"]')).toBeNull();
  });

  it("refuses an empty object: no key of its own at all", () => {
    expect(readPaletteFile("{}")).toBeNull();
  });
});

describe("readPaletteFile: missing and future keys", () => {
  it("fills in missing keys from the shipped defaults", () => {
    const parsed = readPaletteFile('{"primary":"#123456"}');
    expect(parsed).toEqual({ ...DEFAULT_SETTINGS, primary: "#123456" });
  });

  it("reads a file from a future version: unknown extra keys are ignored, not refused", () => {
    const fromTheFuture = JSON.stringify({
      primary: "#654321",
      // Keys this build has never heard of, the way a newer pls,fix's export
      // would look to an older one.
      logoUrl: "https://example.com/logo.png",
      themeVersion: 7,
      gradients: { primary: ["#111111", "#222222"] },
    });
    const parsed = readPaletteFile(fromTheFuture);
    expect(parsed).toEqual({ ...DEFAULT_SETTINGS, primary: "#654321" });
  });

  it("still refuses a future file naming NONE of the current keys", () => {
    const unrelated = JSON.stringify({
      themeVersion: 7,
      gradients: { primary: ["#111111", "#222222"] },
    });
    expect(readPaletteFile(unrelated)).toBeNull();
  });
});

describe("readPaletteFile: a 2 MB file", () => {
  it("refuses 2 MB of garbage without hanging or throwing", () => {
    const garbage = "x".repeat(2 * 1024 * 1024);
    expect(() => readPaletteFile(garbage)).not.toThrow();
    expect(readPaletteFile(garbage)).toBeNull();
  });

  it("still reads a real palette carrying a 2 MB unrelated field", () => {
    const bulky = JSON.stringify({
      primary: "#00FF00",
      // A future export embedding something large this build ignores.
      base64Logo: "A".repeat(2 * 1024 * 1024),
    });
    expect(bulky.length).toBeGreaterThan(2 * 1024 * 1024);
    const parsed = readPaletteFile(bulky);
    expect(parsed).toEqual({ ...DEFAULT_SETTINGS, primary: "#00FF00" });
  });
});
