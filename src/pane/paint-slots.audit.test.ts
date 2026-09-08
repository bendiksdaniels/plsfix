// @vitest-environment jsdom
// The three paintbrush slots as the modeller drives them: the real taskpane
// buttons, through the shared guard, over the strict fake Excel host. Covers
// the wiring nothing else runs - the slot labels, the localStorage mirror, the
// workbook copy, and the boot read that must not undo a capture beside it.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../../test/fakehost";
import type * as ExcelModule from "../excel";
import type * as SharedModule from "./shared";
import type * as SlotsModule from "./paint-slots";

enableStrictLoadSemantics();

const PAINT_KEY = "plsfix.paint.v1";
const PAINT_SETTING = "PLSFIX_PAINT_SLOTS_V1";

let helpers: FakeHelpers;
let smt: typeof ExcelModule;
let shared: typeof SharedModule;
let slots: typeof SlotsModule;
let store: Map<string, string>;

// This Node/jsdom combination ships a global localStorage whose methods are all
// undefined, so the module under test gets a real, controllable one.
function workingStorage(): Pick<Storage, "getItem" | "setItem"> {
  return {
    getItem: (name) => store.get(name) ?? null,
    setItem: (name, value) => {
      store.set(name, value);
    },
  };
}

function throwingStorage(): Pick<Storage, "getItem" | "setItem"> {
  return {
    getItem: () => {
      throw new Error("storage blocked");
    },
    setItem: () => {
      throw new Error("storage blocked");
    },
  };
}

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  store = new Map();
  vi.stubGlobal("localStorage", workingStorage());
  // The real pane markup, so a renamed id fails here instead of in Excel.
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  helpers = installFakeHost({ sheets: ["Model", "Data"] }).helpers;
  smt = await import("../excel");
  shared = await import("./shared");
  slots = await import("./paint-slots");
}

// main.ts's own button loop: the shipped button, through the shared guard.
// A renamed data-action fails here rather than silently doing nothing.
async function click(action: string): Promise<void> {
  const { dispatch } = await import("./dispatch");
  if (!document.querySelector(`button[data-action="${action}"]`)) {
    throw new Error(`taskpane.html has no button for ${action}`);
  }
  await shared.guard(() => dispatch(action), action);
}

function label(index: number): string {
  return document.getElementById(`paint-slot-${index}`)?.textContent ?? "";
}

function toastText(): string {
  return document.getElementById("toast")?.textContent ?? "";
}

// A cell wearing something on every field a slot carries.
function seedSource(): void {
  helpers.seed("Model!A1", [[1]]);
  helpers.setFont("Model!A1", { name: "Georgia", size: 12, bold: true });
  helpers.setNumberFormat("Model!A1", "0.00%");
  helpers.setActiveCell("Model!A1");
  helpers.select("Model!A1");
}

// A slot this workbook saved earlier, in the shape the setting carries.
async function staleWorkbookSlot(): Promise<string> {
  seedSource();
  await click("paint-capture-1");
  const stale = helpers.setting(PAINT_SETTING) ?? "";
  await boot();
  return stale;
}

beforeEach(async () => {
  await boot();
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("capturing a slot from the pane", () => {
  it("labels the slot, mirrors it to storage and saves it with the workbook", async () => {
    seedSource();
    await click("paint-capture-1");

    expect(label(1)).toBe("0.00% · Georgia");
    expect(label(2)).toBe("empty");
    expect(toastText()).toContain("Slot 1: 0.00% · Georgia");
    expect(store.get(PAINT_KEY)).toContain("Georgia");
    expect(helpers.setting(PAINT_SETTING)).toContain("Georgia");
  });

  it("paints the captured look over the next selection", async () => {
    seedSource();
    await click("paint-capture-2");

    helpers.select("Data!B2:C3");
    await click("paint-apply-2");

    expect(toastText()).toContain("Painted slot 2");
    expect(helpers.font("Data!C3").name).toBe("Georgia");
    expect(helpers.numberFormat("Data!C3")).toBe("0.00%");
    // Behind pls,fix Undo like every other write.
    expect(smt.undoTarget()).toBe("Data!B2:C3");
  });

  it("says which slot is empty instead of painting nothing", async () => {
    helpers.select("Model!A1:B2");
    await click("paint-apply-3");

    expect(toastText()).toContain("paintbrush: slot 3 is empty");
  });
});

describe("the slots a machine already had", () => {
  it("renders what localStorage kept", () => {
    store.set(
      PAINT_KEY,
      JSON.stringify([
        null,
        null,
        {
          numberFormat: "#,##0",
          font: {
            name: "Aptos",
            size: 10,
            bold: false,
            italic: false,
            color: "#1F1D1B",
          },
          fill: null,
          horizontalAlignment: "General",
          borders: {
            top: { style: "None", weight: "Thin", color: "#000000" },
            bottom: { style: "None", weight: "Thin", color: "#000000" },
            left: { style: "None", weight: "Thin", color: "#000000" },
            right: { style: "None", weight: "Thin", color: "#000000" },
          },
        },
      ]),
    );

    slots.loadPaintSlots();
    slots.renderPaintSlots();
    expect(label(3)).toBe("#,##0 · Aptos");
  });

  it("reads anything else as three empty slots", async () => {
    store.set(PAINT_KEY, "not json");
    slots.loadPaintSlots();
    slots.renderPaintSlots();

    expect(label(1)).toBe("empty");
    expect(label(3)).toBe("empty");
  });
});

describe("the workbook read at boot", () => {
  it("replaces the machine's slots with the ones the file carries", async () => {
    seedSource();
    await click("paint-capture-1");
    const carried = helpers.setting(PAINT_SETTING);

    // A second pane on another machine: its own localStorage is empty, the
    // workbook is not.
    await boot();
    helpers.setSetting(PAINT_SETTING, carried ?? "");
    await slots.loadWorkbookSlots();

    expect(label(1)).toBe("0.00% · Georgia");
  });

  it("keeps the machine's slots when the workbook never saved any", async () => {
    seedSource();
    await click("paint-capture-1");

    await slots.loadWorkbookSlots();
    expect(label(1)).toBe("0.00% · Georgia");
  });

  // The boot read is started unawaited by main.ts. A capture that lands while
  // it is in flight must survive the older array it brings back, whichever of
  // the two settles first.
  it("never undoes a capture that landed while it was in flight", async () => {
    const stale = await staleWorkbookSlot();
    helpers.setSetting(PAINT_SETTING, stale);
    seedSource();
    helpers.setFont("Model!A1", { name: "Consolas", size: 11, bold: false });

    const inFlight = slots.loadWorkbookSlots();
    await click("paint-capture-1");
    await inFlight;

    expect(label(1)).toBe("0.00% · Consolas");
  });

  it("leaves a capture alone even when the read lands after it", async () => {
    const stale = await staleWorkbookSlot();
    helpers.setSetting(PAINT_SETTING, stale);
    seedSource();
    helpers.setFont("Model!A1", { name: "Consolas", size: 11, bold: false });

    await click("paint-capture-1");
    await slots.loadWorkbookSlots();

    expect(label(1)).toBe("0.00% · Consolas");
  });
});

describe("a webview that refuses local storage", () => {
  it("captures, labels and saves with the workbook all the same", async () => {
    seedSource();
    vi.stubGlobal("localStorage", throwingStorage());
    slots.loadPaintSlots();

    await click("paint-capture-1");
    expect(label(1)).toBe("0.00% · Georgia");
    expect(helpers.setting(PAINT_SETTING)).toContain("Georgia");

    helpers.select("Data!B2");
    await click("paint-apply-1");
    expect(helpers.font("Data!B2").name).toBe("Georgia");
  });

  // Storage refused AND the workbook refusing to keep settings: the capture
  // still has to live for the rest of the session. setItem is the one moment
  // between the capture read and the workbook save, so the refusal is armed
  // from there.
  it("keeps the capture in memory when the workbook refuses too", async () => {
    seedSource();
    vi.stubGlobal("localStorage", {
      getItem: () => null,
      setItem: () => {
        helpers.failNextSync();
        throw new Error("storage blocked");
      },
    });

    await click("paint-capture-1");
    // The refused save is not an error the modeller has to read.
    expect(toastText()).toContain("Slot 1: 0.00% · Georgia");
    expect(label(1)).toBe("0.00% · Georgia");

    helpers.select("Data!B2");
    await click("paint-apply-1");
    expect(helpers.font("Data!B2").name).toBe("Georgia");
  });
});
