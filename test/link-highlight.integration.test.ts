// The linked-cell highlight end to end against the strict fake host: every
// anchored range is tinted, charts and #REF! anchors are left alone, and the
// toggle hands back the fills it covered byte for byte - this session or the
// next time the file is opened.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as ExcelModule from "../src/excel";
import type * as SettingsModule from "../src/settings";
import { anchorName } from "../src/link/model";
import { createWorkspace, type KeyStore } from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
  type FakeWorkbook,
} from "./fakehost";

enableStrictLoadSemantics();

const SETTING = "PLSFIX_LINK_HIGHLIGHT";

let smt: typeof ExcelModule;
let brand: typeof SettingsModule;
let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let relay: FakeRelay;
let tinted: string;

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

async function boot(existing?: FakeWorkbook): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost(
    existing ? { workbook: existing } : { sheets: ["Model", "Data"] },
  );
  helpers = host.helpers;
  workbook = host.workbook;
  smt = await import("../src/excel");
  brand = await import("../src/settings");
  tinted = brand.tint(brand.deriveTheme(brand.DEFAULT_SETTINGS).linkFont, 0.85);
}

// Reopening the file: same workbook model, brand new runtime and module state.
async function reopen(): Promise<void> {
  await boot(workbook);
}

function solid(color: string) {
  return { color, pattern: "Solid", patternColor: color };
}

// Two range links and one chart link, the way the Links tab would leave them.
async function seedLinks(): Promise<{ range: string; chart: string }> {
  const ws = await createWorkspace(memoryStore());
  helpers.seed("Model!B4", [
    [1, 2, 3],
    [4, 5, 6],
  ]);
  helpers.select("Model!B4:D5");
  const first = await smt.exportSelection(ws, relay);

  helpers.seed("Data!A1", [[7, 8]]);
  helpers.select("Data!A1:B1");
  await smt.exportSelection(ws, relay);

  helpers.addChart("Model", {
    name: "Revenue bridge",
    width: 400,
    height: 200,
  });
  helpers.setActiveChart(workbook.charts[0]!);
  const chart = await smt.exportActiveChart(ws, relay);
  return { range: first.id, chart: chart.id };
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
  workbook.fileUrl = "/Users/daniel/Models/Model_v4.xlsx";
  relay = new FakeRelay();
});
afterEach(() => uninstallFakeHost());

describe("toggleLinkHighlight", () => {
  it("tints every anchored range and leaves the chart alone", async () => {
    await seedLinks();
    expect(await smt.toggleLinkHighlight()).toBe(true);

    expect(helpers.fill("Model!B4")).toEqual(solid(tinted));
    expect(helpers.fill("Model!D5")).toEqual(solid(tinted));
    expect(helpers.fill("Data!A1")).toEqual(solid(tinted));
    expect(helpers.fill("Data!B1")).toEqual(solid(tinted));
    // Just outside the linked block, and the chart's own sheet cells.
    expect(helpers.fill("Model!E5").pattern).toBe("None");
    expect(helpers.fill("Model!B6").pattern).toBe("None");
  });

  it("restores the fills it covered byte for byte", async () => {
    await seedLinks();
    helpers.setFill("Model!B4", {
      color: "#EEDDCC",
      pattern: "LightUp",
      patternColor: "#0057B8",
    });
    helpers.setFill("Data!B1", { color: "#123456", pattern: "Solid" });
    const model = helpers.cellMap("Model");
    const data = helpers.cellMap("Data");

    await smt.toggleLinkHighlight();
    expect(helpers.fill("Model!B4").pattern).toBe("Solid");
    expect(helpers.fill("Model!B4").color).toBe(tinted);

    expect(await smt.toggleLinkHighlight()).toBe(false);
    expect(helpers.cellMap("Model")).toEqual(model);
    expect(helpers.cellMap("Data")).toEqual(data);
    expect(helpers.fill("Model!C4").pattern).toBe("None");
  });

  it("stores the snapshot in workbook settings and clears it on toggle-off", async () => {
    await seedLinks();
    await smt.toggleLinkHighlight();

    const stored = JSON.parse(String(helpers.setting(SETTING)));
    expect(stored).toHaveLength(2);
    expect(stored.map((one: { address: string }) => one.address)).toEqual([
      "B4:D5",
      "A1:B1",
    ]);

    await smt.toggleLinkHighlight();
    expect(helpers.setting(SETTING)).toBe("");
  });

  it("skips an anchor whose rows were deleted", async () => {
    const { range } = await seedLinks();
    helpers.breakName(anchorName(range));

    expect(await smt.toggleLinkHighlight()).toBe(true);
    expect(helpers.fill("Model!B4").pattern).toBe("None");
    expect(helpers.fill("Data!A1")).toEqual(solid(tinted));
    const stored = JSON.parse(String(helpers.setting(SETTING)));
    expect(stored).toHaveLength(1);
    expect(stored[0].address).toBe("A1:B1");
  });

  it("refuses linked ranges over the cap without painting anything", async () => {
    const { range } = await seedLinks();
    helpers.setNameFormula(anchorName(range), "=Model!$A$1:$A$5001");

    expect(await rejects(() => smt.toggleLinkHighlight())).toBe(
      "highlight: linked ranges exceed 5,000 cells",
    );
    expect(helpers.fill("Data!A1").pattern).toBe("None");
    expect(helpers.setting(SETTING)).toBeNull();
  });

  it("refuses while the audit overlay owns the fills", async () => {
    await seedLinks();
    helpers.select("Model!B4:D5");
    expect(await smt.toggleAuditOverlay()).toBe(true);
    const painted = helpers.cellMap("Model");

    expect(await rejects(() => smt.toggleLinkHighlight())).toBe(
      "highlight: turn the audit overlay off first",
    );
    // The overlay still owns every fill it painted, and its own snapshot is
    // untouched: toggling it off puts the modeller's formatting back.
    expect(helpers.cellMap("Model")).toEqual(painted);
    expect(helpers.setting(SETTING)).toBeNull();
    expect(await smt.toggleAuditOverlay()).toBe(false);
    expect(helpers.fill("Model!B4").pattern).toBe("None");
  });

  it("says so when this workbook has no range links", async () => {
    expect(await rejects(() => smt.toggleLinkHighlight())).toBe(
      "highlight: no linked ranges in this workbook",
    );
  });
});

// The mirror of the refusal above. Without it the pair is reachable, both
// snapshots are saved in the file, and the two boot restores write over each
// other on the next open.
describe("toggleAuditOverlay against the highlight", () => {
  it("refuses while the highlight owns the fills, and paints nothing", async () => {
    await seedLinks();
    expect(await smt.toggleLinkHighlight()).toBe(true);
    const painted = helpers.cellMap("Model");

    helpers.select("Model!B4:D5");
    expect(await rejects(() => smt.toggleAuditOverlay())).toBe(
      "audit overlay: turn the linked-cell highlight off first",
    );
    expect(helpers.cellMap("Model")).toEqual(painted);
    expect(helpers.setting("smtAuditOverlay")).toBeNull();

    // The highlight is still the sole owner, so its toggle-off is still exact.
    expect(await smt.toggleLinkHighlight()).toBe(false);
    expect(helpers.fill("Model!B4").pattern).toBe("None");
  });

  it("never lets both snapshots be saved in one file", async () => {
    await seedLinks();
    helpers.select("Model!B4:D5");
    expect(await smt.toggleAuditOverlay()).toBe(true);
    await rejects(() => smt.toggleLinkHighlight());
    expect(helpers.setting(SETTING)).toBeNull();

    expect(await smt.toggleAuditOverlay()).toBe(false);
    expect(await smt.toggleLinkHighlight()).toBe(true);
    await rejects(() => smt.toggleAuditOverlay());
    expect(helpers.setting("smtAuditOverlay")).toBe("");
  });
});

// A file written before the refusal was symmetric can still hold both
// snapshots: the highlight painted, the pane was reloaded, and the overlay then
// striped the same cells and snapshotted the tint as if it were the modeller's.
// src/main.ts restores the overlay first and the highlight second, and never
// side by side, because the two orders do not land in the same place.
async function paintBothOverlays(): Promise<Record<string, unknown>> {
  await seedLinks();
  const cell = { formula: "=R[1]C", r1c1: "=R[1]C", value: 1 };
  helpers.seed("Model!B4", [
    [cell, cell, cell],
    [cell, cell, cell],
  ]);
  helpers.setFill("Model!B4", { color: "#EEDDCC", pattern: "Solid" });
  const before = helpers.cellMap("Model");

  await smt.toggleLinkHighlight();
  // The pane restarts: the maps die, the two settings stay in the file, and an
  // older build's overlay is free to paint over the tint it finds.
  await boot(workbook);
  helpers.select("Model!B4:D5");
  await smt.toggleAuditOverlay();
  expect(helpers.setting(SETTING)).not.toBe("");
  expect(helpers.setting("smtAuditOverlay")).not.toBe("");
  return before;
}

// The handover checklist has to see the tint: the modeller ticks the box, runs
// Prepare for sharing, and sends a file with brand-coloured blocks over half
// the grid that the reader has no add-in to take off.
describe("prepare for sharing, with the highlight on", () => {
  it("reports the highlight and the link tokens the workbook carries", async () => {
    await seedLinks();
    const clean = await smt.prepareForSharing();
    expect(clean.report.filter((one) => one.kind === "overlayPainted")).toEqual(
      [],
    );
    expect(clean.report.some((one) => one.kind === "linkTokens")).toBe(true);

    await smt.toggleLinkHighlight();
    const { report } = await smt.prepareForSharing();

    expect(
      report
        .filter((one) => one.kind === "overlayPainted")
        .map((one) => one.label),
    ).toEqual([
      "The linked-cell highlight is still painted over cells; the reader has no add-in to clear it",
    ]);
  });
});

describe("the boot restores, in order", () => {
  it("gives the modeller's own fills back: overlay first, then highlight", async () => {
    const before = await paintBothOverlays();

    await reopen();
    expect(await smt.restorePersistedOverlay()).toBe(true);
    expect(await smt.restoreLinkHighlight()).toBe(true);
    expect(helpers.cellMap("Model")).toEqual(before);
  });

  it("leaves our own tint on the model when the two land the other way round", async () => {
    const before = await paintBothOverlays();

    await reopen();
    expect(await smt.restoreLinkHighlight()).toBe(true);
    expect(await smt.restorePersistedOverlay()).toBe(true);
    expect(helpers.fill("Model!B4").color).toBe(tinted);
    expect(helpers.cellMap("Model")).not.toEqual(before);
  });
});

describe("restoreLinkHighlight", () => {
  it("puts last session's fills back when the file is reopened", async () => {
    await seedLinks();
    helpers.setFill("Model!C5", {
      color: "#EEDDCC",
      pattern: "LightUp",
      patternColor: "#0057B8",
    });
    const before = helpers.cellMap("Model");
    await smt.toggleLinkHighlight();

    // The runtime dies with the pane; the snapshot rides along in the file.
    await reopen();
    expect(await smt.restoreLinkHighlight()).toBe(true);
    expect(helpers.cellMap("Model")).toEqual(before);
    expect(helpers.setting(SETTING)).toBe("");
  });

  it("reports nothing to restore when the highlight was off", async () => {
    expect(await smt.restoreLinkHighlight()).toBe(false);
  });

  it("leaves the audit overlay's own snapshot alone", async () => {
    await seedLinks();
    const formula = { formula: "=R[1]C", r1c1: "=R[1]C", value: 1 };
    helpers.seed("Data!A5", [
      [formula, formula],
      [formula, formula],
    ]);
    const before = helpers.cellMap("Data");
    helpers.select("Data!A5:B6");
    await smt.toggleAuditOverlay();
    const overlaySnapshot = helpers.setting("smtAuditOverlay");
    expect(helpers.fill("Data!A5").pattern).not.toBe("None");

    // A restore for a highlight that never painted must touch neither the
    // overlay's paint nor the snapshot it saved.
    expect(await smt.restoreLinkHighlight()).toBe(false);
    expect(helpers.setting("smtAuditOverlay")).toBe(overlaySnapshot);
    expect(helpers.fill("Data!A5").pattern).not.toBe("None");
    expect(await smt.toggleAuditOverlay()).toBe(false);
    expect(helpers.cellMap("Data")).toEqual(before);
  });
});
