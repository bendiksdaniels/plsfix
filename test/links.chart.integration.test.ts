// "Export active chart" against the strict Excel fake: the picture is exactly
// what it was, and a chart the slide can draw as shapes also ships its data -
// kind, title, categories, values, the source cells' own text as labels and
// brand colours. A chart type, a size or an overlap the slide cannot draw
// leaves the payload the picture it has always been.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import { deriveLinkKeys, open } from "../src/link/crypto";
import {
  decodePayload,
  REGISTRY_SETTING,
  type PicturePayload,
} from "../src/link/model";
import { DEFAULT_SETTINGS } from "../src/settings";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
  type FakeWorkbook,
} from "./fakehost";

enableStrictLoadSemantics();

const CATEGORIES = ["2024A", "2025E", "2026E"];
const VALUES = [1240, 1302, 1400];

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
  links = await import("../src/excel/links");
  // The categories over the values, the way a model lays a chart's block out.
  helpers.seed("Model!B3", [CATEGORIES, VALUES]);
});
afterEach(() => uninstallFakeHost());

function token(): string {
  return JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links[0].token;
}

async function pictureOf(id: string): Promise<PicturePayload> {
  const keys = await deriveLinkKeys(token());
  const stored = relay.links.get(id)!;
  const payload = decodePayload(await open(keys.enc, id, stored.blob));
  if (payload.kind !== "picture") throw new Error(`${id}: not a picture`);
  return payload;
}

describe("exportActiveChart: the chart data beside the picture", () => {
  // The fake's `text` is the value as it was seeded, so the cells wear
  // "General" and their text is the digits alone.
  it("exports a column chart with its data, the cells' text as labels and brand colours", async () => {
    const chart = helpers.addChart("Model", {
      name: "Revenue chart",
      chartType: "ColumnClustered",
      title: "Revenue",
      series: [
        {
          name: "Revenue",
          categories: CATEGORIES,
          values: VALUES,
          valuesSource: "Model!$B$4:$D$4",
        },
      ],
    });
    helpers.setActiveChart(chart);

    const { id } = await links.exportActiveChart(ws, relay);
    const payload = await pictureOf(id);
    expect(payload.png).not.toBe("");
    expect(payload.chart).toMatchObject({
      v: 1,
      kind: "column",
      title: "Revenue",
      categories: CATEGORIES,
      font: DEFAULT_SETTINGS.font,
      titleColor: DEFAULT_SETTINGS.primary,
    });
    expect(payload.chart?.overlap).toBeUndefined();
    expect(payload.chart?.series[0]).toMatchObject({
      name: "Revenue",
      values: VALUES,
      labels: ["1240", "1302", "1400"],
    });
    expect(payload.chart?.series[0]?.colors).toEqual([
      DEFAULT_SETTINGS.accent,
      DEFAULT_SETTINGS.accent,
      DEFAULT_SETTINGS.accent,
    ]);
  });

  it("sends line-chart data beside the picture so PowerPoint can rebuild it", async () => {
    const chart = helpers.addChart("Model", {
      name: "Revenue trend",
      chartType: "Line",
      title: "Revenue",
      series: [{ name: "Revenue", categories: CATEGORIES, values: VALUES }],
    });
    helpers.setActiveChart(chart);

    const { id } = await links.exportActiveChart(ws, relay);
    const payload = await pictureOf(id);
    expect(payload.png).not.toBe("");
    expect(payload.chart).toMatchObject({
      v: 1,
      kind: "line",
      title: "Revenue",
      categories: CATEGORIES,
    });
  });

  // The tornado: one clustered bar chart whose two series sit on the same row.
  it("marks the tornado's overlapped bars", async () => {
    const chart = helpers.addChart("Model", {
      name: "Sensitivity",
      chartType: "BarClustered",
      title: "Sensitivity",
      series: [
        {
          name: "Low",
          categories: CATEGORIES,
          values: [-40, -25, -10],
          overlap: 100,
        },
        {
          name: "High",
          categories: CATEGORIES,
          values: [40, 25, 10],
          overlap: 100,
        },
      ],
    });
    helpers.setActiveChart(chart);

    const { id } = await links.exportActiveChart(ws, relay);
    const payload = await pictureOf(id);
    expect(payload.chart).toMatchObject({ kind: "bar", overlap: true });
    expect(payload.chart?.series.map((one) => one.name)).toEqual([
      "Low",
      "High",
    ]);
    expect(payload.chart?.series[1]?.colors).toEqual([
      DEFAULT_SETTINGS.primary,
      DEFAULT_SETTINGS.primary,
      DEFAULT_SETTINGS.primary,
    ]);
  });

  it("falls back to the picture past forty points", async () => {
    const many = Array.from({ length: 41 }, (_unused, index) =>
      String(index + 1),
    );
    const chart = helpers.addChart("Model", {
      name: "Monthly",
      chartType: "ColumnClustered",
      series: [{ name: "Revenue", categories: many, values: many.map(Number) }],
    });
    helpers.setActiveChart(chart);

    const { id } = await links.exportActiveChart(ws, relay);
    const payload = await pictureOf(id);
    expect(payload.png).not.toBe("");
    expect(payload.chart).toBeUndefined();
  });

  it("formats labels in the house style when the values have no cells", async () => {
    const chart = helpers.addChart("Model", {
      name: "Revenue chart",
      chartType: "ColumnClustered",
      series: [{ name: "Revenue", categories: [], values: VALUES }],
    });
    helpers.setActiveChart(chart);

    const { id } = await links.exportActiveChart(ws, relay);
    const payload = await pictureOf(id);
    // No categories on the series: the slide numbers the points itself.
    expect(payload.chart?.categories).toEqual(["1", "2", "3"]);
    expect(payload.chart?.series[0]?.labels).toEqual([
      "1 240",
      "1 302",
      "1 400",
    ]);
  });
});
