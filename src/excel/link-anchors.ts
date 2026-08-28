// Everything behind the link flows that is not a flow: the registry kept in a
// document setting, the hidden name (or a chart's own name) that keeps pointing
// at a source after rows move under it, the PNG a source renders to, and the
// sealed round trip that puts one record on the relay and in the inbox.

import { deriveLinkKeys, newToken, seal, sha256Hex } from "../link/crypto";
import {
  decodeRegistry,
  encodeInboxItem,
  encodePayload,
  encodeRegistry,
  REGISTRY_SETTING,
  type InboxItem,
  type LinkKind,
  type Payload,
  type Registry,
  type RegistryEntry,
  type Source,
} from "../link/model";
import { base64ToBytes, pngSize } from "../link/png";
import { RelayError, type RelayApi } from "../link/relay";
import type { Workspace } from "../link/workspace";
import { parseAddress } from "./shared";

const CHART_LABEL_SEPARATOR = ": ";
// Charts are laid out in points; rendering at twice that keeps the slide
// picture sharp on a high-density screen.
const CHART_PIXEL_SCALE = 2;

export interface ResolvedRange {
  kind: "range";
  sheet: string;
  ref: string;
  range: Excel.Range;
}

export interface ResolvedChart {
  kind: "chart";
  sheet: string;
  ref: string;
  chart: Excel.Chart;
  width: number;
  height: number;
}

export type ResolvedSource = ResolvedRange | ResolvedChart;

export interface NewLink {
  entry: RegistryEntry;
  src: Source;
  png: string;
  anchor: () => void;
}

// Every failure says which flow it came from and which link; the token is never
// part of a label or a message.
export function staged(stage: string, error: unknown): Error {
  const reason = error instanceof Error ? error.message : String(error);
  return new Error(`${stage}: ${reason}`);
}

// The file name only, never the path: it is shown in PowerPoint and travels in
// every payload. An unsaved workbook has no URL, which reads as "".
export async function workbookName(): Promise<string> {
  return new Promise((done) => {
    const document = Office.context?.document;
    if (!document?.getFilePropertiesAsync) {
      done("");
      return;
    }
    document.getFilePropertiesAsync((result) => {
      const url =
        result.status === Office.AsyncResultStatus.Succeeded
          ? result.value.url
          : "";
      done(url.split(/[\\/]/).pop() ?? "");
    });
  });
}

export async function readRegistry(
  context: Excel.RequestContext,
): Promise<Registry> {
  const setting =
    context.workbook.settings.getItemOrNullObject(REGISTRY_SETTING);
  setting.load("isNullObject,value");
  await context.sync();
  return decodeRegistry(setting.isNullObject ? null : String(setting.value));
}

export function writeRegistry(
  context: Excel.RequestContext,
  registry: Registry,
): void {
  context.workbook.settings.add(REGISTRY_SETTING, encodeRegistry(registry));
}

// A chart entry's label is "<sheet>: <chart name>". The chart itself is called
// the anchor now, so the label holds the only copy of the name it had.
export function chartRef(label: string): string {
  const cut = label.indexOf(CHART_LABEL_SEPARATOR);
  return cut < 0 ? label : label.slice(cut + CHART_LABEL_SEPARATOR.length);
}

export function sourceOf(
  workbook: string,
  anchor: string,
  resolved: ResolvedSource,
): Source {
  return { workbook, sheet: resolved.sheet, ref: resolved.ref, anchor };
}

export function newEntry(
  id: string,
  kind: LinkKind,
  anchor: string,
  label: string,
): RegistryEntry {
  return {
    id,
    kind,
    anchor,
    label,
    token: newToken(),
    createdAt: new Date().toISOString(),
    lastPushedAt: null,
    rev: 0,
  };
}

// Hidden, because the anchor is bookkeeping: it must not clutter the modeller's
// name box, and Excel moves it with its rows the way it moves any name.
export function createRangeAnchor(
  context: Excel.RequestContext,
  range: Excel.Range,
  anchor: string,
): void {
  context.workbook.names.add(anchor, range).visible = false;
}

// A chart cannot carry a defined name, so its own name becomes the anchor.
export function createChartAnchor(chart: Excel.Chart, anchor: string): void {
  chart.name = anchor;
}

async function anchorRange(
  context: Excel.RequestContext,
  anchor: string,
): Promise<Excel.Range | null> {
  const named = context.workbook.names.getItemOrNullObject(anchor);
  named.load("isNullObject");
  await context.sync();
  if (named.isNullObject) return null;

  // The name outlives what it points at: Excel rewrites its formula to #REF!
  // when the rows go, and the name itself stays behind.
  const range = named.getRangeOrNullObject();
  range.load("isNullObject,address,worksheet/name");
  await context.sync();
  return range.isNullObject ? null : range;
}

// A chart keeps its name when it is dragged to another sheet, so every sheet is
// asked - in one batch - rather than trusting the sheet it was exported from.
async function findChart(
  context: Excel.RequestContext,
  anchor: string,
): Promise<Excel.Chart | null> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  const candidates = sheets.items.map((sheet) =>
    sheet.charts.getItemOrNullObject(anchor),
  );
  for (const candidate of candidates) candidate.load("isNullObject");
  await context.sync();

  const found = candidates.find((candidate) => !candidate.isNullObject);
  if (!found) return null;
  found.load("name,width,height,worksheet/name");
  await context.sync();
  return found;
}

// Null means the source is gone: the caller reports it, it never falls back to
// the address the link was created with.
export async function resolveSource(
  context: Excel.RequestContext,
  entry: RegistryEntry,
): Promise<ResolvedSource | null> {
  if (entry.kind === "range") {
    const range = await anchorRange(context, entry.anchor);
    if (!range) return null;
    return {
      kind: "range",
      sheet: range.worksheet.name,
      ref: parseAddress(range.address).address,
      range,
    };
  }
  const chart = await findChart(context, entry.anchor);
  if (!chart) return null;
  return {
    kind: "chart",
    sheet: chart.worksheet.name,
    ref: chartRef(entry.label),
    chart,
    width: chart.width,
    height: chart.height,
  };
}

export async function renderSource(
  context: Excel.RequestContext,
  resolved: ResolvedSource,
): Promise<string> {
  const image =
    resolved.kind === "range"
      ? resolved.range.getImage()
      : resolved.chart.getImage(
          Math.round(resolved.width * CHART_PIXEL_SCALE),
          Math.round(resolved.height * CHART_PIXEL_SCALE),
          Excel.ImageFittingMode.fit,
        );
  await context.sync();
  return image.value;
}

// Removing a link puts the workbook back: the hidden name goes, and a chart
// gets back the name it had before it was anchored.
export async function releaseAnchor(
  context: Excel.RequestContext,
  entry: RegistryEntry,
): Promise<void> {
  if (entry.kind === "range") {
    const named = context.workbook.names.getItemOrNullObject(entry.anchor);
    named.load("isNullObject");
    await context.sync();
    if (!named.isNullObject) named.delete();
    return;
  }
  const chart = await findChart(context, entry.anchor);
  if (chart) chart.name = chartRef(entry.label);
}

export async function pushPayload(
  entry: RegistryEntry,
  src: Source,
  png: string,
  relay: RelayApi,
): Promise<number> {
  const size = pngSize(base64ToBytes(png));
  const payload: Payload = {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: size.width,
    height: size.height,
    png,
    src,
    pushedAt: new Date().toISOString(),
    hash: await sha256Hex(png),
  };
  const keys = await deriveLinkKeys(entry.token);
  const blob = await seal(keys.enc, entry.id, encodePayload(payload));
  return (await relay.putLink(entry.id, keys.auth, blob)).rev;
}

// The note PowerPoint picks up: it carries the token, which is why it is sealed
// with the workspace key the two panes were paired with.
async function announce(
  entry: RegistryEntry,
  src: Source,
  ws: Workspace,
  relay: RelayApi,
): Promise<void> {
  const item: InboxItem = {
    id: entry.id,
    token: entry.token,
    kind: entry.kind,
    label: entry.label,
    src,
    createdAt: entry.createdAt,
  };
  const blob = await seal(ws.enc, ws.id, encodeInboxItem(item));
  await relay.postInbox(ws.id, ws.auth, entry.id, blob);
}

// The workbook is written only once the relay holds the picture, so a failed
// push leaves neither an orphan anchor nor a registry entry behind.
export async function publish(
  context: Excel.RequestContext,
  link: NewLink,
  ws: Workspace,
  relay: RelayApi,
): Promise<void> {
  const { entry, src } = link;
  try {
    entry.rev = await pushPayload(entry, src, link.png, relay);
    entry.lastPushedAt = new Date().toISOString();
    const registry = await readRegistry(context);
    link.anchor();
    registry.links.push(entry);
    writeRegistry(context, registry);
    await context.sync();
    await announce(entry, src, ws, relay);
  } catch (error) {
    throw staged(`export ${entry.label}`, error);
  }
}

// A link the relay never had, or has already dropped, is the outcome we want.
export async function forget(
  entry: RegistryEntry,
  relay: RelayApi,
): Promise<void> {
  const keys = await deriveLinkKeys(entry.token);
  try {
    await relay.deleteLink(entry.id, keys.auth);
  } catch (error) {
    if (!(error instanceof RelayError) || error.kind !== "missing") {
      throw staged(`remove ${entry.label}`, error);
    }
  }
}
