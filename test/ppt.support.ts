// Shared PowerPoint fake-host setup for the ppt.*.integration.test.ts suites:
// the workspace key store, the relay-seeding helpers (seedLink, seedTable,
// seedText and seedChart, pushAgain, pushTable, pushText and pushChart), the
// source-cell constant they publish under, and the per-test boot (bootPpt)
// that installs a fresh fake deck and relay.

import { vi } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import type { ChartData } from "../src/link/chart-model";
import {
  encodePayload,
  newLinkId,
  sourceLabel,
  type InboxItem,
  type Payload,
  type PicturePayload,
  type TableCell,
  type TablePayload,
  type TextPayload,
} from "../src/link/model";
import { base64ToBytes, pngSize } from "../src/link/png";
import type { KeyStore } from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import {
  installFakePpt,
  uninstallFakePpt,
  type FakePresentation,
  type FakePptHelpers,
} from "./fakeppt";
import type * as LinksModule from "../src/ppt/links";

export const src = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

let relay: FakeRelay;

export function memoryStore(): KeyStore {
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

function newId(): string {
  return newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
}

async function publish(
  id: string,
  token: string,
  payload: Payload,
): Promise<void> {
  const keys = await deriveLinkKeys(token);
  await relay.putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
  );
}

function tablePayload(
  cells: TableCell[][],
  widths: number[],
  hash: string,
): TablePayload {
  return {
    v: 1,
    kind: "table",
    rows: cells.length,
    cols: widths.length,
    cells,
    widths,
    src,
    pushedAt: new Date().toISOString(),
    hash,
  };
}

// A table export waiting in the inbox, sealed the way Excel would have sealed
// it: the label is the one sourceLabel gives a table link.
export async function seedTable(
  cells: TableCell[][],
  widths: number[],
): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  const payload = tablePayload(cells, widths, "0".repeat(64));
  await publish(id, token, payload);
  return {
    id,
    token,
    kind: "table",
    label: `${src.sheet}!${src.ref} table`,
    src: payload.src,
    createdAt: payload.pushedAt,
  };
}

export async function pushTable(
  item: InboxItem,
  cells: TableCell[][],
  widths: number[],
): Promise<void> {
  await publish(
    item.id,
    item.token,
    tablePayload(cells, widths, "1".repeat(64)),
  );
}

function textPayload(text: string, hash: string): TextPayload {
  return {
    v: 1,
    kind: "text",
    text,
    src,
    pushedAt: new Date().toISOString(),
    hash,
  };
}

// A text export waiting in the inbox: the label is the one sourceLabel gives
// a text link, so a picture and a text of the same cell stay apart.
export async function seedText(text: string): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  const payload = textPayload(text, "0".repeat(64));
  await publish(id, token, payload);
  return {
    id,
    token,
    kind: "text",
    label: sourceLabel(src, "text"),
    src: payload.src,
    createdAt: payload.pushedAt,
  };
}

export async function pushText(item: InboxItem, text: string): Promise<void> {
  await publish(item.id, item.token, textPayload(text, "1".repeat(64)));
}

export async function seedLink(
  png: string,
  workbook = src.workbook,
): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  const payload: Payload = {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png,
    src: { ...src, workbook },
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };
  await publish(id, token, payload);
  return {
    id,
    token,
    kind: "range",
    label: "Model!B4:F12",
    src: payload.src,
    createdAt: payload.pushedAt,
  };
}

// A column chart of `points` categories: one series, a title and brand
// colours, with the labels Excel already formatted. Its layout is one title,
// one bar and one value label per category, a baseline and one category label,
// so `points` decides the shape count a draw has to spend syncs on.
export function columnChart(points: number): ChartData {
  const categories = Array.from({ length: points }, (_, i) => String(2021 + i));
  const values = categories.map((_, i) => 120 + 20 * i);
  return {
    v: 1,
    kind: "column",
    title: "Revenue",
    categories,
    series: [
      {
        name: "Revenue",
        values,
        labels: values.map(String),
        colors: categories.map(() => "#B27E54"),
      },
    ],
    font: "Aptos Narrow",
    ink: "#282623",
    titleColor: "#14213D",
  };
}

// The picture every host can draw, plus - when the source is a chart type the
// slide knows - the data one with the shape APIs draws natively instead. The
// size in the payload is the picture's own, which is what sets the chart's.
function chartPayload(
  chart: ChartData | null,
  png: string,
  hash: string,
): PicturePayload {
  const size = pngSize(base64ToBytes(png));
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: size.width,
    height: size.height,
    png,
    src,
    pushedAt: new Date().toISOString(),
    hash,
    ...(chart === null ? {} : { chart }),
  };
}

export async function seedChart(
  chart: ChartData,
  png: string,
): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  const payload = chartPayload(chart, png, "0".repeat(64));
  await publish(id, token, payload);
  return {
    id,
    token,
    kind: "chart",
    label: sourceLabel(src, "chart"),
    src: payload.src,
    createdAt: payload.pushedAt,
  };
}

// The next push of a chart link: new data, or none at all when the modeller
// turned the source into a chart type the slide cannot draw.
export async function pushChart(
  item: InboxItem,
  chart: ChartData | null,
  png: string,
): Promise<void> {
  await publish(item.id, item.token, chartPayload(chart, png, "1".repeat(64)));
}

export async function pushAgain(
  item: InboxItem,
  png: string,
  workbook = src.workbook,
): Promise<void> {
  const payload: Payload = {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 1,
    height: 1,
    png,
    src: { ...src, workbook },
    pushedAt: new Date().toISOString(),
    hash: "1".repeat(64),
  };
  await publish(item.id, item.token, payload);
}

// The per-test boot: a fresh fake deck (3 slides, slide 0 selected) and a
// fresh relay, the same sequence every ppt.*.integration.test.ts beforeEach
// used to run inline.
export async function bootPpt(): Promise<{
  links: typeof LinksModule;
  presentation: FakePresentation;
  helpers: FakePptHelpers;
  relay: FakeRelay;
}> {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 3 });
  const presentation = host.presentation;
  const helpers = host.helpers;
  relay = new FakeRelay();
  const links = await import("../src/ppt/links");
  helpers.selectSlide(presentation.slides[0]!.id);
  return { links, presentation, helpers, relay };
}
