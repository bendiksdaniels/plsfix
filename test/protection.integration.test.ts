// A protected sheet. Excel refuses every write to a locked cell with a bare
// AccessDenied, which reaches the pane as an error over something the modeller
// only asked to have painted. Autocolor and the audit overlay ask first and
// report a line instead; nothing on the sheet changes either way.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  smt = await import("../src/excel");
}

function seedGrid(): void {
  helpers.seed("Model!A1", [
    [1, { formula: "=A1*2", value: 2 }],
    [3, { formula: "=A2*2", value: 6 }],
  ]);
  helpers.select("Model!A1:B2");
}

beforeEach(async () => {
  await boot();
});

describe("autocolor on a protected sheet", () => {
  it("skips the paint and says so", async () => {
    seedGrid();
    helpers.protectSheet("Model");
    const before = helpers.cellMap("Model");

    expect(await smt.autocolorSelection()).toBe(
      "Autocolor: this sheet is protected, nothing was changed",
    );
    expect(helpers.cellMap("Model")).toEqual(before);
    // Nothing was overwritten, so nothing was captured either.
    expect(smt.undoTarget()).toBeNull();
  });

  it("paints and counts the cells on an unprotected one", async () => {
    seedGrid();
    expect(await smt.autocolorSelection()).toBe("Autocolor: 4 cells");
  });

  it("skips the colour key too", async () => {
    helpers.setActiveCell("Model!D2");
    helpers.protectSheet("Model");

    expect(await smt.insertColorKey()).toBe(
      "Color key: this sheet is protected, nothing was changed",
    );
    expect(helpers.value("Model!D2")).toBe("");
  });

  it("writes the colour key on an unprotected sheet", async () => {
    helpers.setActiveCell("Model!D2");
    expect(await smt.insertColorKey()).toBe("Color key added");
    expect(helpers.value("Model!D2")).toBe("Color key");
  });
});

describe("the audit overlay on a protected sheet", () => {
  it("stays off and leaves a note behind", async () => {
    seedGrid();
    helpers.protectSheet("Model");
    const before = helpers.cellMap("Model");

    expect(await smt.toggleAuditOverlay()).toBe(false);
    // Bare, with no stage of its own: the pane's toggleAudit() composes the
    // "Audit overlay on/off" state onto it, so the stage names once, not
    // twice (src/excel/protection.ts's protectedSentence).
    expect(smt.lastAuditNote()).toBe(
      "this sheet is protected, nothing was changed",
    );
    // Read once: a later toggle must not inherit it.
    expect(smt.lastAuditNote()).toBeNull();
    expect(helpers.cellMap("Model")).toEqual(before);
    expect(smt.auditOverlayOn()).toBe(false);
  });

  it("paints and reports nothing on an unprotected sheet", async () => {
    seedGrid();
    expect(await smt.toggleAuditOverlay()).toBe(true);
    expect(smt.lastAuditNote()).toBeNull();
  });
});

describe("an edit on a protected sheet", () => {
  it("fails by name rather than with Excel's own string", async () => {
    seedGrid();
    helpers.protectSheet("Model");

    await expect(smt.applySignFlip()).rejects.toThrow(
      "Sign flip: this sheet is protected, nothing was changed",
    );
    await expect(smt.applyNumberFormat("whole")).rejects.toThrow(
      "Number formatting: this sheet is protected, nothing was changed",
    );
    expect(helpers.value("Model!A1")).toBe(1);
  });
});

describe("locked cells inside an unprotected island", () => {
  it("drops the paint rather than the pane", async () => {
    seedGrid();
    // The modeller may type in B1:B2 only; the rest of the sheet is locked, and
    // autocolor writes runs that cross the boundary.
    helpers.protectSheet("Model", ["Model!B1:B2"]);

    expect(await smt.autocolorSelection()).toBe(
      "Autocolor: this sheet is protected, nothing was changed",
    );
    expect(helpers.font("Model!B1").color).toBe("#000000");
  });
});
