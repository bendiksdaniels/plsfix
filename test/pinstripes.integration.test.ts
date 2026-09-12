// Pinstripes against the strict fake host: every second row (or column) of the
// selection banded in the brand tint, the second press clearing them again, the
// two overlays it refuses to paint over, and every selection a modeller can
// hand it - multi-area, merged, protected, over the cap and empty.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import { createWorkspace, type KeyStore } from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import { DEFAULT_SETTINGS, tint } from "../src/settings";

enableStrictLoadSemantics();

// 10 % of the palette's primary over white, the band tint the flow paints.
const BAND = tint(DEFAULT_SETTINGS.primary, 0.9);

let helpers: FakeHelpers;
let smt: typeof ExcelModule;
let relay: FakeRelay;

function memoryStore(): KeyStore {
  const map = new Map<string, string>();
  return {
    get: async (k) => map.get(k) ?? null,
    set: async (k, v) => {
      map.set(k, v);
    },
    remove: async (k) => {
      map.delete(k);
    },
  };
}

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  smt = await import("../src/excel");
  relay = new FakeRelay();
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

// A five-row block of numbers at A1:C5.
function seedBlock(): void {
  helpers.seed(
    "Model!A1",
    Array.from({ length: 5 }, (_unused, row) => [row, row + 1, row + 2]),
  );
  helpers.select("Model!A1:C5");
}

beforeEach(async () => {
  await boot();
});

describe("the band tint", () => {
  it("is the palette's primary at 10 %, channel by channel", () => {
    const channels = [1, 3, 5].map((at) =>
      Number.parseInt(DEFAULT_SETTINGS.primary.slice(at, at + 2), 16),
    );
    const expected = channels
      .map((c) =>
        Math.round(255 - (255 - c) * 0.1)
          .toString(16)
          .padStart(2, "0"),
      )
      .join("");
    expect(BAND).toBe(`#${expected.toUpperCase()}`);
  });
});

describe("applyPinstripes over rows", () => {
  it("bands every second row and leaves the odd ones alone", async () => {
    seedBlock();

    expect(await smt.applyPinstripes("rows")).toBe("Pinstripes: 2 rows banded");

    expect(helpers.fill("Model!A2").color).toBe(BAND);
    expect(helpers.fill("Model!C2").color).toBe(BAND);
    expect(helpers.fill("Model!A4").color).toBe(BAND);
    expect(helpers.fill("Model!A1").pattern).toBe("None");
    expect(helpers.fill("Model!A3").pattern).toBe("None");
    expect(helpers.fill("Model!A5").pattern).toBe("None");
  });

  it("clears the bands on a second press", async () => {
    seedBlock();
    await smt.applyPinstripes("rows");

    expect(await smt.applyPinstripes("rows")).toBe(
      "Pinstripes: 2 rows cleared",
    );
    expect(helpers.fill("Model!A2").pattern).toBe("None");
    expect(helpers.fill("Model!A4").pattern).toBe("None");
  });

  it("bands again when one banded row was painted over by hand", async () => {
    seedBlock();
    await smt.applyPinstripes("rows");
    helpers.setFill("Model!A4", { color: "#FFFF00", pattern: "Solid" });

    expect(await smt.applyPinstripes("rows")).toBe("Pinstripes: 2 rows banded");
    expect(helpers.fill("Model!A4").color).toBe(BAND);
  });

  it("puts the modeller's own fills back on Undo", async () => {
    seedBlock();
    helpers.setFill("Model!A2", { color: "#FFFF00", pattern: "Solid" });
    await smt.applyPinstripes("rows");
    expect(helpers.fill("Model!A2").color).toBe(BAND);

    await smt.undoLastAction();
    expect(helpers.fill("Model!A2").color).toBe("#FFFF00");
  });
});

describe("applyPinstripes over columns", () => {
  it("bands every second column", async () => {
    seedBlock();

    expect(await smt.applyPinstripes("columns")).toBe(
      "Pinstripes: 1 column banded",
    );
    expect(helpers.fill("Model!B1").color).toBe(BAND);
    expect(helpers.fill("Model!B5").color).toBe(BAND);
    expect(helpers.fill("Model!A1").pattern).toBe("None");
    expect(helpers.fill("Model!C1").pattern).toBe("None");
  });

  it("clears them on a second press", async () => {
    seedBlock();
    await smt.applyPinstripes("columns");

    expect(await smt.applyPinstripes("columns")).toBe(
      "Pinstripes: 1 column cleared",
    );
    expect(helpers.fill("Model!B1").pattern).toBe("None");
  });
});

describe("what Pinstripes refuses", () => {
  it("refuses a selection with only one row", async () => {
    helpers.seed("Model!A1", [[1, 2, 3]]);
    helpers.select("Model!A1:C1");

    expect(await rejects(() => smt.applyPinstripes("rows"))).toBe(
      "Pinstripes need at least two rows in the selection.",
    );
  });

  it("refuses a selection with only one column", async () => {
    seedBlock();
    helpers.select("Model!A1:A5");

    expect(await rejects(() => smt.applyPinstripes("columns"))).toBe(
      "Pinstripes need at least two columns in the selection.",
    );
  });

  it("refuses a ctrl-clicked selection", async () => {
    seedBlock();
    helpers.selectAreas(["Model!A1:C5", "Model!E1:F2"]);

    expect(await rejects(() => smt.applyPinstripes("rows"))).toBe(
      "Pinstripes: select a single range",
    );
  });

  it("refuses a selection over the cell cap, before any fill is read", async () => {
    helpers.select("Model!A:C");

    expect(await rejects(() => smt.applyPinstripes("rows"))).toBe(
      "Pinstripes supports up to 5,000 selected cells at once.",
    );
  });

  it("says the sheet is protected and paints nothing", async () => {
    seedBlock();
    helpers.protectSheet("Model");

    expect(await smt.applyPinstripes("rows")).toBe(
      "Pinstripes: this sheet is protected, nothing was changed",
    );
    expect(helpers.fill("Model!A2").pattern).toBe("None");
  });

  it("refuses while the audit overlay owns the fills", async () => {
    helpers.seed("Model!B4", [
      [
        { formula: "=A4+1", r1c1: "=RC[-1]+1" },
        { formula: "=B4+1", r1c1: "=RC[-1]+1" },
      ],
      [
        { formula: "=A5+1", r1c1: "=RC[-1]+1" },
        { formula: "=B5+1", r1c1: "=RC[-1]+1" },
      ],
    ]);
    helpers.select("Model!B4:C5");
    expect(await smt.toggleAuditOverlay()).toBe(true);
    const painted = helpers.cellMap("Model");

    helpers.select("Model!A1:C5");
    expect(await rejects(() => smt.applyPinstripes("rows"))).toBe(
      "Pinstripes: turn the audit overlay off first",
    );
    expect(helpers.cellMap("Model")).toEqual(painted);
  });

  it("refuses while the linked-cell highlight owns the fills", async () => {
    const ws = await createWorkspace(memoryStore());
    helpers.seed("Data!A1", [[7, 8]]);
    helpers.select("Data!A1:B1");
    await smt.exportSelection(ws, relay);
    expect(await smt.toggleLinkHighlight()).toBe(true);
    const painted = helpers.cellMap("Data");

    seedBlock();
    expect(await rejects(() => smt.applyPinstripes("rows"))).toBe(
      "Pinstripes: turn the linked-cell highlight off first",
    );
    expect(helpers.cellMap("Data")).toEqual(painted);
    expect(helpers.fill("Model!A2").pattern).toBe("None");
  });
});

describe("Pinstripes on an empty and a merged selection", () => {
  it("bands empty cells the way the fill cycle does", async () => {
    helpers.select("Model!C3:E6");

    expect(await smt.applyPinstripes("rows")).toBe("Pinstripes: 2 rows banded");
    expect(helpers.fill("Model!C4").color).toBe(BAND);
    expect(helpers.fill("Model!C6").color).toBe(BAND);
  });

  it("bands a merged band with the row it sits in", async () => {
    helpers.seed("Model!A1", [["Heading"]]);
    helpers.merge("Model!A2:B2");
    helpers.seed("Model!A3", [
      [1, 2],
      [3, 4],
    ]);
    helpers.select("Model!A1:B4");

    expect(await smt.applyPinstripes("rows")).toBe("Pinstripes: 2 rows banded");
    expect(helpers.fill("Model!A2").color).toBe(BAND);
    expect(helpers.fill("Model!B2").color).toBe(BAND);
    expect(helpers.fill("Model!A4").color).toBe(BAND);
  });
});
