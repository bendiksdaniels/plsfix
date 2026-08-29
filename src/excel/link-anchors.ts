// The registry kept in a document setting, and the anchor that keeps pointing
// at a source after rows move under it: a hidden name, or a chart's own name.
// Owns registry read/write, workbookName, anchor create/resolve/delete and
// render. The relay round trip that ships a render lives in link-record.ts.

import { deriveLinkKeys, newToken } from "../link/crypto";
import {
  ANCHOR_PREFIX,
  emptyRegistry,
  encodeRegistry,
  REGISTRY_SETTING,
  tryDecodeRegistry,
  type LinkKind,
  type Registry,
  type RegistryEntry,
  type Source,
} from "../link/model";
import { isRelayError, type RelayApi } from "../link/relay";
import { hostSupports } from "./internal";
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

// Every failure says which flow it came from and which link; the token is never
// part of a label or a message.
export function staged(stage: string, error: unknown): Error {
  const reason = error instanceof Error ? error.message : String(error);
  return new Error(`${stage}: ${reason}`);
}

// Range.getImage and the chart image surface both arrived in ExcelApi 1.9;
// without them there is no picture to send, so a flow stops before it anchors
// anything.
export function requireImageApi(): void {
  if (!hostSupports("1.9")) {
    throw new Error("export: Excel 2021 / Microsoft 365 required");
  }
}

export function entryOf(
  registry: Registry,
  id: string,
  stage: string,
): RegistryEntry {
  const entry = registry.links.find((link) => link.id === id);
  if (!entry) throw new Error(`${stage} ${id}: not in this workbook`);
  return entry;
}

// A chart's own name is its anchor, so exporting an anchored chart a second
// time would rename it and silently orphan the first link. An anchor name no
// registry entry claims is the orphan itself - left behind by an export that
// failed after the rename - and re-anchoring it is the only way back.
export function refuseAnchoredChart(registry: Registry, name: string): void {
  if (!name.startsWith(ANCHOR_PREFIX)) return;
  const existing = registry.links.find((link) => link.anchor === name);
  if (!existing) return;
  throw new Error(
    `export chart: already linked as ${existing.label}; push it instead`,
  );
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
  if (setting.isNullObject) return emptyRegistry();

  const raw = String(setting.value);
  const registry = tryDecodeRegistry(raw);
  // No setting is a workbook with no links yet. A setting we cannot read is
  // someone else's data or a newer schema, and every flow here writes the
  // whole setting back: decoding it to empty would erase every link record.
  if (registry === null && raw.trim() !== "") {
    throw new Error("registry PLSFIX_LINKS: unreadable, not overwriting");
  }
  return registry ?? emptyRegistry();
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
): Excel.NamedItem {
  const named = context.workbook.names.add(anchor, range);
  named.visible = false;
  return named;
}

// A chart cannot carry a defined name, so its own name becomes the anchor.
export function createChartAnchor(chart: Excel.Chart, anchor: string): void {
  chart.name = anchor;
}

// One entry's handles, in the order the phases fill them in: the range a hidden
// name points at, or the chart carrying the anchor as its own name.
interface Resolving {
  entry: RegistryEntry;
  named: Excel.NamedItem | null;
  range: Excel.Range | null;
  candidates: Excel.Chart[];
  chart: Excel.Chart | null;
}

// A chart keeps its name when it is dragged to another sheet, so every sheet is
// asked rather than trusting the sheet the link was exported from.
function chartCandidates(
  sheets: Excel.WorksheetCollection | null,
  anchor: string,
): Excel.Chart[] {
  if (!sheets) return [];
  return sheets.items.map((sheet) => sheet.charts.getItemOrNullObject(anchor));
}

// Phase 2: the name outlives what it points at - Excel rewrites its formula to
// #REF! when the rows go and leaves the name behind - so the range is asked for
// separately, and every entry's question rides in the same batch.
function queueTargets(
  resolving: Resolving[],
  sheets: Excel.WorksheetCollection | null,
): void {
  for (const one of resolving) {
    if (one.entry.kind === "range") {
      if (!one.named || one.named.isNullObject) continue;
      one.range = one.named.getRangeOrNullObject();
      one.range.load("isNullObject,address,worksheet/name");
      continue;
    }
    one.candidates = chartCandidates(sheets, one.entry.anchor);
    for (const candidate of one.candidates) candidate.load("isNullObject");
  }
}

function assemble(one: Resolving): ResolvedSource | null {
  if (one.entry.kind === "range") {
    const { range } = one;
    if (!range || range.isNullObject) return null;
    return {
      kind: "range",
      sheet: range.worksheet.name,
      ref: parseAddress(range.address).address,
      range,
    };
  }
  const { chart } = one;
  if (!chart) return null;
  return {
    kind: "chart",
    sheet: chart.worksheet.name,
    ref: chartRef(one.entry.label),
    chart,
    width: chart.width,
    height: chart.height,
  };
}

// Every anchor resolved in three syncs whatever the link count: the names and
// the sheet list, then each name's range and each sheet's chart of that name,
// then the geometry of the charts that answered. Auto-push runs this on a timer
// and the highlight over every anchored range, so one round trip per link would
// cost a workbook of them every time the typing pauses. Null in a slot means
// that source is gone: the caller reports it, and it never falls back to the
// address the link was created with.
export async function resolveSources(
  context: Excel.RequestContext,
  entries: RegistryEntry[],
): Promise<(ResolvedSource | null)[]> {
  if (entries.length === 0) return [];
  const resolving: Resolving[] = entries.map((entry) => ({
    entry,
    named:
      entry.kind === "range"
        ? context.workbook.names.getItemOrNullObject(entry.anchor)
        : null,
    range: null,
    candidates: [],
    chart: null,
  }));
  for (const one of resolving) one.named?.load("isNullObject");

  const wantsCharts = entries.some((entry) => entry.kind === "chart");
  const sheets = wantsCharts ? context.workbook.worksheets : null;
  sheets?.load("items/name");
  await context.sync();

  queueTargets(resolving, sheets);
  await context.sync();

  for (const one of resolving) {
    one.chart = one.candidates.find((chart) => !chart.isNullObject) ?? null;
    one.chart?.load("name,width,height,worksheet/name");
  }
  if (resolving.some((one) => one.chart !== null)) await context.sync();

  return resolving.map(assemble);
}

export async function resolveSource(
  context: Excel.RequestContext,
  entry: RegistryEntry,
): Promise<ResolvedSource | null> {
  const [resolved] = await resolveSources(context, [entry]);
  return resolved ?? null;
}

// The anchor is bound in the same batch the picture is asked for, and office.js
// batches are not transactional: the rename or the names.add executes and only
// the getImage fails, so the workbook is already mutated when the sync rejects.
// A failed render therefore undoes its own anchor rather than leaving one no
// registry entry claims.
export async function renderAnchored(
  context: Excel.RequestContext,
  resolved: ResolvedSource,
  label: string,
  release: () => void,
): Promise<string> {
  try {
    return await renderSource(context, resolved);
  } catch (error) {
    await undoAnchor(context, release);
    throw staged(`export ${label}`, error);
  }
}

// Best effort, like publish's rollback: if the workbook will not take the undo,
// the render error the caller is about to see is the one worth reporting.
async function undoAnchor(
  context: Excel.RequestContext,
  release: () => void,
): Promise<void> {
  try {
    release();
    await context.sync();
  } catch {
    return;
  }
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
  const resolved = await resolveSource(context, entry);
  if (resolved?.kind === "chart") resolved.chart.name = chartRef(entry.label);
}

// A link the relay never had, or has already dropped, is the outcome we want.
// Anything else means the relay is still serving that picture to everyone
// holding the deck, so the caller keeps its record: the message says the
// workbook was left alone and the revoke can be tried again.
export async function forget(
  entry: RegistryEntry,
  relay: RelayApi,
): Promise<void> {
  const keys = await deriveLinkKeys(entry.token);
  try {
    await relay.deleteLink(entry.id, keys.auth);
  } catch (error) {
    if (isRelayError(error) && error.kind === "missing") return;
    const reason = error instanceof Error ? error.message : String(error);
    throw new Error(
      `remove ${entry.label}: ${reason}; nothing was removed, try again`,
    );
  }
}
