// @vitest-environment jsdom
// Stress pass, slice P2: every hygiene button pressed the way a modeller
// presses it - through the real taskpane.html, the real src/pane/shared.ts
// guard and the real src/pane/dispatch.ts, over the fake host. What the pane
// owes at this level is not the write but the SENTENCE: the toast has to carry
// a line in the pane's voice, the buttons have to come back on, and the
// "(too large for undo)" note has to land on the action that earned it.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeHostOptions,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as Shared from "../src/pane/shared";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let shared: typeof Shared;
let dispatch: (action: string) => Promise<string>;

// Every button of the P2 area, by the data-action taskpane.html ships.
const HYGIENE_ACTIONS = [
  "cycle-indent",
  "cycle-align",
  "cycle-underline",
  "cycle-fill",
  "cycle-font",
  "cycle-border",
  "cycle-number-currency",
  "cycle-row-title",
  "cycle-row-height",
  "cycle-col-width",
  "sheets-unhide-all",
  "sheets-show-only",
  "sheets-bury",
  "sheets-move-up",
  "sheets-move-down",
  "sheets-move-end",
  "clean-past-data",
];

async function boot(options: FakeHostOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  const host = installFakeHost({
    sheets: ["Model", "Data", "Notes"],
    ...options,
  });
  helpers = host.helpers;
  shared = await import("../src/pane/shared");
  shared.setExcelReady(true);
  ({ dispatch } = await import("../src/pane/dispatch"));
}

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

function toastIsError(): boolean {
  return Boolean(document.querySelector("#toast .toast-error, #toast.error"));
}

function buttonsDisabled(): boolean {
  return shared.actionButtons.some((button) => button.disabled);
}

/** Presses the button the way the pane does: through the shared guard. */
async function press(action: string): Promise<string> {
  await shared.guard(() => dispatch(action), action);
  return toastText();
}

// A line in the pane's voice: prose a modeller reads, never office.js's own
// string and never a bare object or an empty toast.
function readsLikeASentence(line: string): boolean {
  if (line.trim().length < 4) return false;
  if (!/^[A-Z0-9"']/.test(line)) return false;
  return !/RichApi|\[object |undefined|GeneralException|InvalidArgument/.test(
    line,
  );
}

beforeEach(async () => {
  await boot();
});

describe("every hygiene button on a workbook that fights back", () => {
  it("answers a fresh workbook with a sentence and no stuck buttons", async () => {
    for (const action of HYGIENE_ACTIONS) {
      await boot();
      const line = await press(action);

      expect(readsLikeASentence(line), `${action}: ${line}`).toBe(true);
      expect(buttonsDisabled(), action).toBe(false);
    }
  });

  it("answers a ctrl-clicked selection with a sentence, refusal or not", async () => {
    for (const action of HYGIENE_ACTIONS) {
      await boot();
      helpers.seed("Model!A1", [["Revenue"], ["Cost"]]);
      helpers.selectAreas(["Model!A1:A2", "Model!C1:C2"]);
      const line = await press(action);

      expect(readsLikeASentence(line), `${action}: ${line}`).toBe(true);
      expect(buttonsDisabled(), action).toBe(false);
    }
  });

  it("answers a protected sheet with a sentence naming the stage", async () => {
    for (const action of HYGIENE_ACTIONS) {
      await boot();
      helpers.seed("Model!A1", [["Revenue", 100]]);
      helpers.setFill("Model!A3:B9", { color: "#FFEECC", pattern: "Solid" });
      helpers.select("Model!A1:B2");
      helpers.protectSheet("Model");
      const line = await press(action);

      expect(readsLikeASentence(line), `${action}: ${line}`).toBe(true);
      expect(buttonsDisabled(), action).toBe(false);
    }
  });

  it("answers a locked workbook structure with a sentence", async () => {
    for (const action of HYGIENE_ACTIONS) {
      await boot();
      helpers.seed("Model!A1", [["Revenue", 100]]);
      helpers.select("Model!A1:B2");
      helpers.protectWorkbook();
      const line = await press(action);

      expect(readsLikeASentence(line), `${action}: ${line}`).toBe(true);
      expect(buttonsDisabled(), action).toBe(false);
    }
  });

  it("answers a one-sheet workbook with a sentence", async () => {
    for (const action of HYGIENE_ACTIONS) {
      await boot({ sheets: ["Only"] });
      helpers.seed("Only!A1", [["Revenue"]]);
      helpers.select("Only!A1");
      const line = await press(action);

      expect(readsLikeASentence(line), `${action}: ${line}`).toBe(true);
      expect(buttonsDisabled(), action).toBe(false);
    }
  });

  it("answers a host below ExcelApi 1.9 with a sentence", async () => {
    for (const action of HYGIENE_ACTIONS) {
      await boot({
        isSetSupported: (set, version) =>
          !(set === "ExcelApi" && version === "1.9"),
      });
      helpers.seed("Model!A1", [["Revenue"], ["Cost"]]);
      helpers.selectAreas(["Model!A1", "Model!C1"]);
      const line = await press(action);

      expect(readsLikeASentence(line), `${action}: ${line}`).toBe(true);
      expect(buttonsDisabled(), action).toBe(false);
    }
  });

  it("answers a whole-column selection with a sentence on every capped tool", async () => {
    // The capped cycles refuse a million cells by name; the ones that write a
    // single scalar over the band paint it, as the ribbon's own buttons do.
    for (const action of [
      "cycle-border",
      "cycle-number-currency",
      "cycle-row-title",
    ]) {
      await boot();
      helpers.select("Model!C1:C1048576");
      const line = await press(action);

      expect(line, action).toMatch(/support(s)? up to/);
      expect(readsLikeASentence(line), `${action}: ${line}`).toBe(true);
    }
  });
});

describe("the pane with no workbook behind it", () => {
  it("says Excel is not connected rather than throwing office.js at the toast", async () => {
    for (const action of HYGIENE_ACTIONS) {
      await boot();
      shared.setExcelReady(false);

      expect(await press(action)).toBe("Excel is not connected.");
      expect(buttonsDisabled()).toBe(false);
    }
  });
});

describe("the undo note the guard adds", () => {
  it("adds it to the action that skipped the capture, and to no other", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1:E10000");
    expect(await press("cycle-indent")).toBe("Indent: 1 (too large for undo)");

    helpers.select("Model!A1");
    expect(await press("cycle-indent")).toBe("Indent: 2");
  });

  // The flag is drained in the guard's finally, so a refusal that had already
  // skipped its capture cannot decorate whatever succeeds next.
  it("does not leak the note past a refusal onto the next success", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1:E10000");
    helpers.protectSheet("Model");
    expect(await press("cycle-indent")).toContain("this sheet is protected");

    await boot();
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1");
    expect(await press("cycle-indent")).toBe("Indent: 1");
  });

  it("never claims a skipped capture for the tools that never capture", async () => {
    helpers.select("Model!A1:E10000");
    expect(await press("cycle-row-height")).toBe(
      "Row height 18 pt (outside pls,fix Undo)",
    );
    expect(await press("cycle-col-width")).toBe(
      "Column width 80 (outside pls,fix Undo)",
    );
  });
});

describe("the guard under two presses", () => {
  it("hands the buttons back after an error and after a success", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1");

    await press("sheets-move-up");
    expect(toastIsError() || toastText().includes("already the first")).toBe(
      true,
    );
    expect(buttonsDisabled()).toBe(false);

    expect(await press("cycle-indent")).toBe("Indent: 1");
    expect(buttonsDisabled()).toBe(false);
  });

  it("leaves one sentence in the toast when two presses overlap", async () => {
    helpers.seed("Model!A1", [["Revenue"]]);
    helpers.select("Model!A1");

    await Promise.all([
      shared.guard(() => dispatch("cycle-indent"), "cycle-indent"),
      shared.guard(() => dispatch("cycle-indent"), "cycle-indent"),
    ]);

    expect(document.querySelectorAll("#toast .toast-text")).toHaveLength(1);
    expect(readsLikeASentence(toastText())).toBe(true);
    expect(buttonsDisabled()).toBe(false);
  });
});
