// The Templates flow against the strict fake host: where a block lands, what it
// writes into each cell, the brand look it wears, the refusal over anything
// already standing there, and the Undo that takes it all back.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import type * as SettingsModule from "../src/settings";
import { cellAddress } from "../src/find";
import { TEMPLATES, templateById } from "../src/templates";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;
let brand: typeof SettingsModule;
let palette: (typeof SettingsModule)["DEFAULT_SETTINGS"];
let theme: ReturnType<(typeof SettingsModule)["deriveTheme"]>;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
  brand = await import("../src/settings");
  palette = brand.DEFAULT_SETTINGS;
  theme = brand.deriveTheme(palette);
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

beforeEach(async () => {
  await boot();
});

// ---------------------------------------------------------------------------

describe("insertTemplate", () => {
  it("writes the DCF block at the active cell", async () => {
    helpers.select("Model!C5");

    expect(await smt.insertTemplate("dcf")).toBe(
      "Template written: DCF valuation (13x6)",
    );

    expect(helpers.value("Model!C5")).toBe("DCF valuation");
    expect(helpers.value("Model!C6")).toBe("WACC");
    expect(helpers.value("Model!D6")).toBe(0.09);
    expect(helpers.value("Model!D9")).toBe(1);
    expect(helpers.value("Model!H9")).toBe(5);
    expect(helpers.value("Model!D10")).toBe(100_000);
  });

  it("resolves every formula against the corner it landed on", async () => {
    helpers.select("Model!C5");
    await smt.insertTemplate("dcf");

    expect(helpers.formula("Model!D11")).toBe("=1/(1+D6)^D9");
    expect(helpers.formula("Model!D12")).toBe("=D10*D11");
    expect(helpers.formula("Model!D14")).toBe("=SUM(D12:H12)");
    expect(helpers.formula("Model!D15")).toBe("=H10*(1+D7)/(D6-D7)");
    expect(helpers.formula("Model!D16")).toBe("=D15*H11");
    expect(helpers.formula("Model!D17")).toBe("=D14+D16");
  });

  it("moves every reference when the same block lands elsewhere", async () => {
    helpers.select("Model!A1");
    await smt.insertTemplate("dcf");

    expect(helpers.formula("Model!B7")).toBe("=1/(1+B2)^B5");
    expect(helpers.formula("Model!B13")).toBe("=B10+B12");
  });

  it("dresses the block in the brand presets", async () => {
    helpers.select("Model!C5");
    await smt.insertTemplate("dcf");

    // Title row.
    expect(helpers.fill("Model!C5").color).toBe(theme.titleFill);
    expect(helpers.font("Model!C5")).toMatchObject({
      color: theme.titleText,
      bold: true,
      size: 15,
      name: palette.font,
    });
    expect(helpers.rowHeight("Model", 4)).toBe(25);

    // Header row, with the rule under it.
    expect(helpers.fill("Model!D9").color).toBe(theme.headerFill);
    expect(helpers.font("Model!D9").bold).toBe(true);
    expect(helpers.border("Model!D9", "bottom")).toEqual({
      style: "Continuous",
      color: theme.headerBorder,
      weight: "Thin",
    });

    // Hardcoded inputs blue, formulas in the formula colour.
    expect(helpers.font("Model!D6").color).toBe(theme.inputFont);
    expect(helpers.font("Model!D10").color).toBe(theme.inputFont);
    expect(helpers.font("Model!D11").color).toBe(theme.formulaFont);
    expect(helpers.font("Model!C6").color).toBe(theme.formulaFont);
    expect(helpers.fill("Model!D11").pattern).toBe("None");

    // The answer row.
    expect(helpers.fill("Model!C17").color).toBe(theme.resultFill);
    expect(helpers.font("Model!D17").bold).toBe(true);
    expect(helpers.border("Model!C17", "top")).toMatchObject({
      style: "Double",
      color: theme.resultBorder,
    });
  });

  it("takes its number formats from the pane's own", async () => {
    helpers.select("Model!C5");
    await smt.insertTemplate("dcf");

    const currency = brand.currencyNumberFormat(
      palette.currency,
      palette.language,
    );
    expect(helpers.numberFormat("Model!D6")).toBe("0.0%;[Red](0.0%);-");
    expect(helpers.numberFormat("Model!D9")).toBe("#,##0;[Red](#,##0);-");
    expect(helpers.numberFormat("Model!D10")).toBe(currency);
    expect(helpers.numberFormat("Model!D17")).toBe(currency);
    expect(helpers.numberFormat("Model!C5")).toBe("General");
  });

  it("selects the block it wrote", async () => {
    helpers.select("Model!C5");
    await smt.insertTemplate("dcf");

    expect(workbook.selectionAddress()).toBe("Model!C5:H17");
  });

  it("refuses a target block with anything standing in it", async () => {
    helpers.seed("Model!E10", [["Base case"]]);
    helpers.select("Model!C5");

    expect(await rejects(() => smt.insertTemplate("dcf"))).toBe(
      "Templates need an empty block of 13 rows by 6 columns at the selection.",
    );
    expect(helpers.value("Model!C5")).toBe("");
    expect(helpers.value("Model!E10")).toBe("Base case");
  });

  it("refuses a template it does not know", async () => {
    expect(await rejects(() => smt.insertTemplate("mystery"))).toBe(
      "Unknown template: mystery",
    );
  });

  it("restores the sheet exactly on undo", async () => {
    helpers.select("Model!C5");
    const before = helpers.cellMap("Model");

    await smt.insertTemplate("debt-schedule");
    expect(helpers.cellMap("Model")).not.toEqual(before);

    expect(await smt.undoLastAction()).toBe("Model!C5:H24");
    expect(helpers.cellMap("Model")).toEqual(before);
  });
});

// ---------------------------------------------------------------------------

describe("every template", () => {
  it("lands whole at the active cell, titled by its own name", async () => {
    for (const template of TEMPLATES) {
      await boot();
      helpers.select("Model!B3");

      expect(await smt.insertTemplate(template.id)).toBe(
        `Template written: ${template.name} (${template.rows}x${template.cols})`,
      );
      expect(helpers.value("Model!B3")).toBe(template.name);

      const corner = cellAddress(2 + template.rows - 1, 1 + template.cols - 1);
      expect(workbook.selectionAddress()).toBe(`Model!B3:${corner}`);

      template.cells.forEach((row, rowIndex) => {
        row.forEach((cell, columnIndex) => {
          const at = `Model!${cellAddress(2 + rowIndex, 1 + columnIndex)}`;
          const written = helpers.formula(at);
          if (cell.f !== undefined) {
            expect(String(written).startsWith("=")).toBe(true);
            expect(String(written)).not.toContain("{");
          } else {
            expect(written).toBe(cell.v ?? "");
          }
        });
      });
    }
  });
});

// ---------------------------------------------------------------------------

describe("the EBITDA bridge template", () => {
  it("is the shape the waterfall tool reads", async () => {
    const bridge = templateById("ebitda-bridge");
    expect(bridge).not.toBeNull();

    helpers.select("Model!A1");
    await smt.insertTemplate("ebitda-bridge");

    // The seven bridge rows only: the title above them names the chart and the
    // check row below stays out of it.
    helpers.select("Model!A2:B8");
    expect(await smt.insertWaterfall()).toContain("7 points, ties at");
    expect(workbook.charts[0]?.title).toBe("EBITDA bridge");
  });

  it("checks back to zero against the closing total", async () => {
    helpers.select("Model!A1");
    await smt.insertTemplate("ebitda-bridge");

    expect(helpers.value("Model!A9")).toBe("Check");
    expect(helpers.formula("Model!B9")).toBe("=B2+SUM(B3:B7)-B8");
  });
});
