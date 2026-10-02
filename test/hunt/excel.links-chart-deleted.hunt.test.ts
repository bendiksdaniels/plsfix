// Attacks the "chart was deleted" edge across links.ts + link-anchors.ts: a
// chart link whose chart no longer exists anywhere in the workbook (not
// merely moved to another sheet, which links.integration.test already
// covers) - simulated as no chart of that name on any sheet.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../../src/excel/links";
import { anchorName } from "../../src/link/model";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../../src/link/workspace";
import { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
  type FakeWorkbook,
} from "../fakehost";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let relay: FakeRelay;
let ws: Workspace;

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

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  workbook = host.workbook;
  workbook.fileUrl = "/Users/daniel/Models/Model_v4.xlsx";
  relay = new FakeRelay();
  ws = await createWorkspace(memoryStore());
  links = await import("../../src/excel/links");
});
afterEach(() => uninstallFakeHost());

async function exportChart(): Promise<string> {
  helpers.addChart("Model", {
    name: "Revenue bridge",
    width: 400,
    height: 200,
  });
  helpers.setActiveChart(workbook.charts[0]!);
  const result = await links.exportActiveChart(ws, relay);
  return result.id;
}

// The anchor rename already committed by export; deleting the chart record
// afterwards is exactly what the workbook looks like once the modeller
// deletes that chart - no sheet's charts collection has it any more.
function deleteChart(id: string): void {
  const index = workbook.charts.findIndex(
    (chart) => chart.name === anchorName(id),
  );
  expect(index).toBeGreaterThanOrEqual(0);
  workbook.charts.splice(index, 1);
}

describe("a chart link whose chart was deleted", () => {
  it("lists as missing instead of throwing", async () => {
    const id = await exportChart();
    deleteChart(id);

    const rows = await links.listWorkbookLinks();

    expect(rows).toHaveLength(1);
    expect(rows[0]).toMatchObject({ source: "missing" });
  });

  it("counts as missing on a push, never as failed, and touches no other link", async () => {
    const chartId = await exportChart();
    helpers.seed("Data!B2", [[1, 2]]);
    helpers.select("Data!B2:C2");
    const { id: rangeId } = await links.exportSelection(ws, relay);
    deleteChart(chartId);

    const summary = await links.pushLinks("all", relay);

    expect(summary).toEqual({ pushed: 1, missing: 1, failed: 0, failures: [] });
    expect(relay.links.get(rangeId)!.rev).toBe(2);
    // The deleted chart's own relay copy is untouched: nothing was pushed
    // for it, so its last good picture is exactly what a deck still shows.
    expect(relay.links.get(chartId)!.rev).toBe(1);
  });

  it("keeps reporting missing on three repeated pushes, not just the first", async () => {
    const id = await exportChart();
    deleteChart(id);

    for (let attempt = 0; attempt < 3; attempt += 1) {
      const summary = await links.pushLinks("all", relay);
      expect(summary).toEqual({
        pushed: 0,
        missing: 1,
        failed: 0,
        failures: [],
      });
    }
  });

  it("refuses to jump to the source with a plain sentence", async () => {
    const id = await exportChart();
    deleteChart(id);

    await expect(links.goToSource(id)).rejects.toThrow(
      /go to source Model: Revenue bridge: source missing/,
    );
  });

  // A second chart of the very same anchor NAME re-added on another sheet
  // (Excel numbers new charts, but nothing stops a modeller renaming one to
  // match by hand) must not be mistaken for the deleted one: the link keeps
  // pointing at whichever chart currently carries that exact name, which is
  // the documented anchor rule (a chart's own name IS the anchor) - not a
  // bug, but worth pinning so a future change to chartCandidates cannot
  // silently start preferring the "wrong" sheet.
  it("re-attaches to a same-named chart added back after the deletion", async () => {
    const id = await exportChart();
    deleteChart(id);
    helpers.addChart("Data", { name: anchorName(id), width: 200, height: 100 });

    const rows = await links.listWorkbookLinks();

    expect(rows[0]).toMatchObject({ source: "ok" });
    const summary = await links.pushLinks("all", relay);
    expect(summary).toEqual({ pushed: 1, missing: 0, failed: 0, failures: [] });
  });
});
