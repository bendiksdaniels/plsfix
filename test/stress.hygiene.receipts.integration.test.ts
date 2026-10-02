// @vitest-environment jsdom
// Slice W part 2: every format cycle now answers with the step it landed on
// instead of the generic "Selection updated" - proven here by walking each
// cycle a full lap through the real dispatch() over the fake host, and
// checking every step against src/cycle-labels.ts, the one place the wording
// lives. A label table that drifts from its cycle table (src/cycles.ts) fails
// here, not just in a human review of a diff.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import {
  BORDER_CYCLE_LABELS,
  FILL_CYCLE_LABELS,
  FONT_CYCLE_LABELS,
  numberCycleLabels,
  ROW_STYLE_LABELS,
} from "../src/cycle-labels";
import {
  buildNumberCycles,
  buildRowStyleCycles,
  type NumberCycleFamily,
  type RowStyleKind,
} from "../src/cycles";
import { DEFAULT_SETTINGS } from "../src/settings";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let dispatch: (action: string) => Promise<string>;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  const shared = await import("../src/pane/shared");
  shared.setExcelReady(true);
  ({ dispatch } = await import("../src/pane/dispatch"));
}

beforeEach(async () => {
  await boot();
  helpers.seed("Model!A1", [["Revenue"]]);
  helpers.select("Model!A1");
});

// A fresh cell's numberFormat/fill/font never match a cycle's own steps, so
// the first press always lands on step 0 - the same "-1 steps to 0" rule
// src/cycles.ts documents for every cycle here.
describe("number format cycles", () => {
  const families = Object.keys(
    buildNumberCycles(DEFAULT_SETTINGS),
  ) as NumberCycleFamily[];

  it.each(families)(
    "walks the whole %s cycle and back to step 0",
    async (family) => {
      const labels = numberCycleLabels(DEFAULT_SETTINGS)[family];
      for (const label of labels) {
        expect(await dispatch(`cycle-number-${family}`)).toBe(
          `Number format: ${label}`,
        );
      }
      // One more press: the cycle wraps, not stalls on the last step.
      expect(await dispatch(`cycle-number-${family}`)).toBe(
        `Number format: ${labels[0]}`,
      );
    },
  );
});

describe("row style cycles", () => {
  const kinds = Object.keys(
    buildRowStyleCycles(DEFAULT_SETTINGS),
  ) as RowStyleKind[];

  it.each(kinds)(
    "walks the whole %s row style and back to step 0",
    async (kind) => {
      for (const label of ROW_STYLE_LABELS[kind]) {
        expect(await dispatch(`cycle-row-${kind}`)).toBe(`Row style: ${label}`);
      }
      expect(await dispatch(`cycle-row-${kind}`)).toBe(
        `Row style: ${ROW_STYLE_LABELS[kind][0]}`,
      );
    },
  );
});

it("walks the whole fill cycle and back to step 0", async () => {
  for (const label of FILL_CYCLE_LABELS) {
    expect(await dispatch("cycle-fill")).toBe(`Fill: ${label}`);
  }
  expect(await dispatch("cycle-fill")).toBe(`Fill: ${FILL_CYCLE_LABELS[0]}`);
});

it("walks the whole font colour cycle and back to step 0", async () => {
  for (const label of FONT_CYCLE_LABELS) {
    expect(await dispatch("cycle-font")).toBe(`Font colour: ${label}`);
  }
  expect(await dispatch("cycle-font")).toBe(
    `Font colour: ${FONT_CYCLE_LABELS[0]}`,
  );
});

it("walks the whole border cycle and back to step 0", async () => {
  // A 2x2 block, not the single A1 cell every other cycle here uses: box and
  // grid draw the same four outline edges and differ only in the inside
  // lines a single cell has none of, so on A1 alone they read back as the
  // same state and the cycle would skip grid entirely.
  helpers.seed("Model!A1", [
    ["Revenue", 100],
    ["Cost", 40],
  ]);
  helpers.select("Model!A1:B2");
  // Unlike the other cycles, "none" is a real, matchable rung (index 0) and
  // a fresh block already reads as it - so the first press moves past it:
  // seven presses land on steps 1,2,3,4,5,0,1 in that order.
  for (const step of [1, 2, 3, 4, 5, 0, 1]) {
    expect(await dispatch("cycle-border")).toBe(
      `Borders: ${BORDER_CYCLE_LABELS[step]}`,
    );
  }
});

describe("hygiene cycles", () => {
  it("walks the whole indent ladder and back to 0", async () => {
    expect(await dispatch("cycle-indent")).toBe("Indent: 1");
    expect(await dispatch("cycle-indent")).toBe("Indent: 2");
    expect(await dispatch("cycle-indent")).toBe("Indent: 3");
    expect(await dispatch("cycle-indent")).toBe("Indent: 0");
  });

  it("walks the whole alignment ladder in the brief's own words", async () => {
    expect(await dispatch("cycle-align")).toBe("Aligned: left");
    expect(await dispatch("cycle-align")).toBe("Aligned: centre");
    expect(await dispatch("cycle-align")).toBe("Aligned: right");
    expect(await dispatch("cycle-align")).toBe("Aligned: general");
    expect(await dispatch("cycle-align")).toBe("Aligned: left");
  });

  it("walks the whole underline ladder in the brief's own words", async () => {
    expect(await dispatch("cycle-underline")).toBe("Underline: single");
    expect(await dispatch("cycle-underline")).toBe("Underline: double");
    expect(await dispatch("cycle-underline")).toBe("Underline: none");
    expect(await dispatch("cycle-underline")).toBe("Underline: single");
  });
});

describe("row height and column width, outside pls,fix Undo", () => {
  it("names the point size it landed the row on", async () => {
    expect(await dispatch("cycle-row-height")).toBe(
      "Row height 18 pt (outside pls,fix Undo)",
    );
    expect(await dispatch("cycle-row-height")).toBe(
      "Row height 21 pt (outside pls,fix Undo)",
    );
  });

  it("names the width it landed the column on, with no unit", async () => {
    expect(await dispatch("cycle-col-width")).toBe(
      "Column width 80 (outside pls,fix Undo)",
    );
    expect(await dispatch("cycle-col-width")).toBe(
      "Column width 96 (outside pls,fix Undo)",
    );
  });
});

describe("the IFERROR guard", () => {
  it("counts what it added, then counts what it stripped back off", async () => {
    helpers.seed("Model!A1", [["=B1/C1", "=B2*2"]]);
    helpers.select("Model!A1:B1");
    expect(await dispatch("if-error")).toBe("IFERROR added to 2 formulas");
    expect(await dispatch("if-error")).toBe("IFERROR stripped from 2 formulas");
  });

  it("names both counts when a mixed selection changes both ways", async () => {
    helpers.seed("Model!A1", [["=IFERROR(B1/C1,0)", "=B2*2"]]);
    helpers.select("Model!A1:B1");
    expect(await dispatch("if-error")).toBe(
      "IFERROR added to 1 formula, stripped from 1 formula",
    );
  });
});

it("names the sheet and cell CAGR landed at", async () => {
  helpers.seed("Model!A1", [[100], [110], [121]]);
  helpers.select("Model!A1:A3");
  expect(await dispatch("cagr")).toBe("CAGR written at Model!A4");
});
