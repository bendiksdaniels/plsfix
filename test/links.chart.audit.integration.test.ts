// Audit suite for the Excel half of a chart link: the hosts and refusals that
// leave the payload the picture it has always been, and the sentence the
// toast shows for each - the one both panes share. Strict load semantics are
// on, so a batch that reads a scalar it never loaded fails here.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import {
  CHARTS_NEED_EXCEL_1_12,
  SERIES_READ_REFUSED,
} from "../src/excel/link-chart";
import { chartCapIssue, pictureNote } from "../src/link/chart-model";
import { deriveLinkKeys, open } from "../src/link/crypto";
import {
  decodePayload,
  REGISTRY_SETTING,
  type PicturePayload,
} from "../src/link/model";
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
  helpers.seed("Model!B3", [CATEGORIES, VALUES]);
});
afterEach(() => uninstallFakeHost());

async function pictureOf(id: string): Promise<PicturePayload> {
  const token = JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links[0]
    .token as string;
  const keys = await deriveLinkKeys(token);
  const stored = relay.links.get(id)!;
  const payload = decodePayload(await open(keys.enc, id, stored.blob));
  if (payload.kind !== "picture") throw new Error(`${id}: not a picture`);
  return payload;
}

function addColumnChart(valuesSource?: string): void {
  const chart = helpers.addChart("Model", {
    name: "Revenue chart",
    chartType: "ColumnClustered",
    title: "Revenue",
    series: [
      {
        name: "Revenue",
        categories: CATEGORIES,
        values: VALUES,
        ...(valuesSource === undefined ? {} : { valuesSource }),
      },
    ],
  });
  helpers.setActiveChart(chart);
}

describe("a host that cannot describe a chart at all", () => {
  it("ships the picture and names the Excel build below getDimensionValues", async () => {
    // ExcelApi 1.12 is what reads a series' categories and values; below it
    // there is nothing to send but the picture, and the toast says so.
    helpers.setSupported((_set, version) => version !== "1.12");
    addColumnChart("Model!$B$4:$D$4");

    const result = await links.exportActiveChart(ws, relay);
    const payload = await pictureOf(result.id);
    expect(payload.png).not.toBe("");
    expect(payload.chart).toBeUndefined();
    expect(payload.chartIssue).toBe(CHARTS_NEED_EXCEL_1_12);
    expect(result.note).toBe(pictureNote(CHARTS_NEED_EXCEL_1_12));
  });

  it("says Excel refused the read when the series will not answer twice", async () => {
    addColumnChart("Model!$B$4:$D$4");
    // The head read rides the picture's own batch; the plain retry asks the
    // same chart again, and a chart office.js will not describe refuses both.
    helpers.failNextChartRead();
    helpers.failNextChartRead();

    const result = await links.exportActiveChart(ws, relay);
    const payload = await pictureOf(result.id);
    // The picture is what the deck gets, with the reason beside it.
    expect(payload.png).not.toBe("");
    expect(payload.chart).toBeUndefined();
    expect(payload.chartIssue).toBe(SERIES_READ_REFUSED);
    expect(result.note).toBe(pictureNote(SERIES_READ_REFUSED));
  });
});

describe("the labels a chart's own cells give", () => {
  it("falls back to the house style when the source sheet has gone", async () => {
    // The series still answers, but the range it names is on a sheet the
    // workbook no longer has: the numbers are formatted the house way rather
    // than costing the link its chart data.
    addColumnChart("Ghost!$B$4:$D$4");

    const { id } = await links.exportActiveChart(ws, relay);
    const payload = await pictureOf(id);
    expect(payload.chart?.series[0]?.labels).toEqual([
      "1 240",
      "1 302",
      "1 400",
    ]);
  });

  it("uses the cells' own text when the source range is there", async () => {
    addColumnChart("Model!$B$4:$D$4");

    const { id } = await links.exportActiveChart(ws, relay);
    const payload = await pictureOf(id);
    expect(payload.chart?.series[0]?.labels).toEqual(["1240", "1302", "1400"]);
  });
});

describe("the cap sentences the two panes share", () => {
  it("gives the Excel toast the sentence chartCapIssue writes", async () => {
    const chart = helpers.addChart("Model", {
      name: "Seven",
      chartType: "ColumnClustered",
      title: "Seven",
      series: Array.from({ length: 7 }, (_unused, index) => ({
        name: `S${String(index)}`,
        categories: CATEGORIES,
        values: VALUES,
      })),
    });
    helpers.setActiveChart(chart);

    const result = await links.exportActiveChart(ws, relay);
    const payload = await pictureOf(result.id);
    // Not a string of its own: the toast, the payload's reason and the note
    // the PowerPoint pane shows are all this one sentence.
    const issue = chartCapIssue("column", CATEGORIES.length, 7)!;
    expect(payload.chartIssue).toBe(issue);
    expect(result.note).toBe(pictureNote(issue));
  });
});
