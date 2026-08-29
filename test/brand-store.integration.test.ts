// The workbook-level brand store against the strict fake host: what the pane
// writes comes back after the file is reopened, a workbook that never carried a
// palette reads as null, and the setting read goes through load() + sync().

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import { DEFAULT_SETTINGS, serializeSettings } from "../src/settings";

// A scalar read with no load() plus context.sync() behind it throws here, just
// as it does in Excel.
enableStrictLoadSemantics();

const BRAND_SETTING = "plsfix.brand.v1";

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

// A fresh runtime plus a fresh module graph: nothing the pane cached survives.
async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
}

// Reopening the file: same workbook model, brand new runtime and module state.
async function reopen(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ workbook });
  helpers = host.helpers;
  smt = await import("../src/excel");
}

function palette(patch: Partial<typeof DEFAULT_SETTINGS>): string {
  return serializeSettings({ ...DEFAULT_SETTINGS, ...patch });
}

beforeEach(boot);

describe("workbook brand store", () => {
  it("reads null from a workbook that never carried a palette", async () => {
    expect(await smt.readWorkbookBrand()).toBeNull();
  });

  it("round trips the palette JSON the pane serializes", async () => {
    const json = palette({ primary: "#123456", accent: "#654321" });
    await smt.writeWorkbookBrand(json);

    expect(helpers.setting(BRAND_SETTING)).toBe(json);
    expect(await smt.readWorkbookBrand()).toBe(json);
  });

  it("hands the palette back after the workbook is reopened", async () => {
    const json = palette({ accent: "#00FF00", font: "Calibri" });
    await smt.writeWorkbookBrand(json);

    // The pane's runtime dies with the window; the setting rides in the file,
    // which is the whole point: another computer opens the same brand.
    await reopen();
    expect(await smt.readWorkbookBrand()).toBe(json);
  });

  it("replaces the stored palette instead of keeping both", async () => {
    await smt.writeWorkbookBrand(palette({}));
    const next = palette({ currency: "$", autocolorOnEdit: true });
    await smt.writeWorkbookBrand(next);

    expect(await smt.readWorkbookBrand()).toBe(next);
    expect(workbook.settings.size).toBe(1);
  });

  it("treats a blank setting as no palette of its own", async () => {
    helpers.setSetting(BRAND_SETTING, "");
    expect(await smt.readWorkbookBrand()).toBeNull();
  });

  it("leaves malformed JSON to the caller's parse step", async () => {
    helpers.setSetting(BRAND_SETTING, "{not json");
    expect(await smt.readWorkbookBrand()).toBe("{not json");
  });
});

// Without this the green suite above would only prove that the fake happens to
// answer settings reads from its own model.
describe("strict load discipline", () => {
  interface SettingProbe {
    value: string;
    isNullObject: boolean;
  }

  async function readWithoutLoad(): Promise<string> {
    const host = globalThis as unknown as {
      Excel: { run: (cb: (context: never) => unknown) => Promise<unknown> };
    };
    let thrown = "nothing was thrown";
    await host.Excel.run(async (context: never) => {
      const settings = (
        context as unknown as {
          workbook: {
            settings: {
              getItemOrNullObject: (key: string) => SettingProbe;
            };
          };
        }
      ).workbook.settings;
      const probe = settings.getItemOrNullObject(BRAND_SETTING);
      try {
        await Promise.resolve(probe.value);
      } catch (error) {
        thrown = (error as Error).message;
      }
    });
    return thrown;
  }

  it("refuses a setting read with no load behind it", async () => {
    await smt.writeWorkbookBrand(palette({}));
    expect(await readWithoutLoad()).toContain(
      "The property 'value' is not available",
    );
  });
});
