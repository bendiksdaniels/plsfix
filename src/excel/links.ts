// The Excel side of tracked links, one function per pane action: export the
// active chart, push every anchored source again, list what this workbook
// owns, jump back to a source and remove a link. The selection exports live in
// link-export.ts and are re-exported here. The anchor - not the address it was
// created at - is what every later flow resolves through.

import { randomBytes } from "../link/crypto";
import {
  anchorName,
  newLinkId,
  sourceLabel,
  type RegistryEntry,
} from "../link/model";
import { pictureNote } from "../link/chart-model";
import { isRelayError, type RelayApi } from "../link/relay";
import type { Workspace } from "../link/workspace";
import { hostSupports } from "./internal";
import {
  createChartAnchor,
  entryOf,
  forget,
  newEntry,
  readRegistry,
  refuseAnchoredChart,
  releaseAnchor,
  requireImageApi,
  resolveSource,
  resolveSources,
  sourceOf,
  workbookName,
  writeRegistry,
  type ResolvedSource,
} from "./link-anchors";
import type { ExportResult } from "./link-export";
import { renderAnchored, renderSource } from "./link-render";
import { exclusive } from "./link-lock";
import { announce, publish, pushPayload, type NewLink } from "./link-record";

// The one relay refusal a modeller can act on: the sealed export is past the
// relay's body limit (413), so a smaller range is the way out.
const TOO_BIG_TO_SEND =
  "That export is too big to send. Export a smaller range.";

export { workbookName };
export * from "./link-export";

export interface WorkbookLinkRow {
  entry: RegistryEntry;
  source: "ok" | "missing";
}

export interface PushSummary {
  pushed: number;
  missing: number;
  failed: number;
  // One "<label>: <reason>" per failed push, so the pane can say why rather
  // than only how many.
  failures: string[];
}

export interface PushOptions {
  // Local mode: also post each pushed link's inbox row, sealed with this
  // workspace, so a deck that does not hold the link yet can insert it.
  announce?: Workspace;
}

// The same guard formatSelectedChart uses: the hosted office.js always defines
// the method, so the host's API set is what decides.
function requireChartApi(context: Excel.RequestContext): void {
  const callable = (context.workbook as unknown as Record<string, unknown>)
    .getActiveChartOrNullObject;
  if (typeof callable !== "function") {
    throw new Error("Exporting a chart needs a newer Excel build.");
  }
}

// The chart to export: the selected one, else the one the pane picked, else
// the sheet's only chart. Excel on the web cannot select a chart by name, so
// the pick is what makes chart export possible there at all.
async function chartToExport(
  context: Excel.RequestContext,
  pick: string | null,
): Promise<Excel.Chart> {
  requireChartApi(context);
  const active = context.workbook.getActiveChartOrNullObject();
  active.load("isNullObject");
  await context.sync();
  const chart = active.isNullObject
    ? await chartOnActiveSheet(context, pick)
    : active;
  chart.load("name,width,height,worksheet/name");
  await context.sync();
  return chart;
}

async function chartOnActiveSheet(
  context: Excel.RequestContext,
  pick: string | null,
): Promise<Excel.Chart> {
  const charts = context.workbook.worksheets.getActiveWorksheet().charts;
  charts.load("items/name");
  await context.sync();
  const names = charts.items.map((chart) => chart.name);
  const name = pick ?? (names.length === 1 ? (names[0] ?? null) : null);
  if (name === null) {
    throw new Error(
      names.length === 0
        ? "No chart on this sheet."
        : "Select a chart first, or pick one from the list.",
    );
  }
  const chart = charts.getItemOrNullObject(name);
  chart.load("isNullObject");
  await context.sync();
  if (chart.isNullObject)
    throw new Error(`No chart named ${name} on this sheet.`);
  return chart;
}

// Read-only: the names the Links tab offers in its chart list.
export async function listActiveSheetCharts(): Promise<string[]> {
  return Excel.run(async (context) => {
    const charts = context.workbook.worksheets.getActiveWorksheet().charts;
    charts.load("items/name");
    await context.sync();
    return charts.items.map((chart) => chart.name);
  });
}

export async function exportActiveChart(
  ws: Workspace,
  relay: RelayApi,
  pick: string | null = null,
): Promise<ExportResult> {
  requireImageApi();
  const workbook = await workbookName();
  return exclusive("export chart", () =>
    Excel.run(async (context) => {
      const registry = await readRegistry(context);
      const chart = await chartToExport(context, pick);
      const previousName = chart.name;
      refuseAnchoredChart(registry, previousName);

      const resolved: ResolvedSource = {
        kind: "chart",
        sheet: chart.worksheet.name,
        ref: previousName,
        chart,
        width: chart.width,
        height: chart.height,
      };
      const id = newLinkId(randomBytes);
      const anchor = anchorName(id);
      const src = sourceOf(workbook, anchor, resolved);
      const entry = newEntry(
        id,
        "chart",
        anchor,
        sourceLabel(src, "chart"),
        registry.activeProject,
      );
      // The rename commits with the batch that asks for the picture, so a render
      // that fails has to give the chart its own name back: an PLSFIX_LINK_ chart no
      // registry entry claims is one the modeller cannot export again.
      createChartAnchor(chart, anchor);
      const release = () => createChartAnchor(chart, previousName);
      const render = await renderAnchored(
        context,
        resolved,
        entry.label,
        release,
      );

      const link: NewLink = { entry, src, render, registry, release };
      await publish(context, link, ws, relay);
      const note =
        render.kind === "picture" && render.chartIssue !== undefined
          ? pictureNote(render.chartIssue)
          : undefined;
      return {
        id,
        label: entry.label,
        ...(note === undefined ? {} : { note }),
      };
    }),
  );
}

// Read-only, so it stays out of the queue: a list drawn while a push is in
// flight is one refresh behind, never a registry written back over one.
export async function listWorkbookLinks(): Promise<WorkbookLinkRow[]> {
  return Excel.run(async (context) => {
    const registry = await readRegistry(context);
    const resolved = await resolveSources(context, registry.links);
    return registry.links.map((entry, index): WorkbookLinkRow => ({
      entry,
      source: resolved[index] ? "ok" : "missing",
    }));
  });
}

export async function pushLinks(
  ids: string[] | "all",
  relay: RelayApi,
  options: PushOptions = {},
): Promise<PushSummary> {
  return exclusive("push", () => pushRegistry(ids, relay, options));
}

// A push is a report, not an assertion: a source that is gone and a relay that
// refused are counted rather than thrown, so one bad link cannot stop the rest.
// The caller must already hold the link queue - auto-push holds it across
// deciding which links an edit touched and pushing them, which is one section.
export async function pushRegistry(
  ids: string[] | "all",
  relay: RelayApi,
  options: PushOptions = {},
): Promise<PushSummary> {
  const workbook = await workbookName();
  return Excel.run(async (context) => {
    const registry = await readRegistry(context);
    const wanted = ids === "all" ? null : new Set(ids);
    if (wanted) for (const id of wanted) entryOf(registry, id, "push");

    const summary: PushSummary = {
      pushed: 0,
      missing: 0,
      failed: 0,
      failures: [],
    };
    const targets = registry.links.filter(
      (entry) => !wanted || wanted.has(entry.id),
    );
    const sources = await resolveSources(context, targets);
    for (const [index, entry] of targets.entries()) {
      const resolved = sources[index] ?? null;
      if (resolved === null) {
        summary.missing += 1;
        continue;
      }
      await pushOne(
        context,
        entry,
        resolved,
        workbook,
        relay,
        summary,
        options,
      );
    }
    writeRegistry(context, registry);
    await context.sync();
    return summary;
  });
}

// Counted, never thrown: the entry is updated in place, so the registry the
// caller writes back carries the new revision.
async function pushOne(
  context: Excel.RequestContext,
  entry: RegistryEntry,
  resolved: ResolvedSource,
  workbook: string,
  relay: RelayApi,
  summary: PushSummary,
  options: PushOptions,
): Promise<void> {
  try {
    const render = await renderSource(context, resolved);
    const src = sourceOf(workbook, entry.anchor, resolved);
    entry.rev = await pushPayload(entry, src, render, relay);
    entry.lastPushedAt = new Date().toISOString();
    if (options.announce) await announce(entry, src, options.announce, relay);
    summary.pushed += 1;
  } catch (error) {
    // A 413 is the one relay refusal a modeller can act on: the sealed
    // export is past the relay's body limit, so a smaller range is the way.
    const reason =
      isRelayError(error) && error.kind === "tooLarge"
        ? TOO_BIG_TO_SEND
        : error instanceof Error
          ? error.message
          : String(error);
    summary.failed += 1;
    summary.failures.push(`${entry.label}: ${reason}`);
  }
}

// Anchors are searched on every worksheet, hidden ones included, but Excel
// refuses to activate a sheet that is hidden: a linked range on a hidden calc
// sheet would fail the sync with a bare InvalidOperation.
async function requireVisibleSheet(
  context: Excel.RequestContext,
  resolved: ResolvedSource,
  label: string,
): Promise<void> {
  const sheet =
    resolved.kind === "chart"
      ? resolved.chart.worksheet
      : resolved.range.worksheet;
  sheet.load("name,visibility");
  await context.sync();
  if (sheet.visibility !== Excel.SheetVisibility.visible) {
    throw new Error(`go to source ${label}: sheet "${sheet.name}" is hidden`);
  }
}

export async function goToSource(id: string): Promise<void> {
  await Excel.run(async (context) => {
    const registry = await readRegistry(context);
    const entry = entryOf(registry, id, "go to source");
    const resolved = await resolveSource(context, entry);
    if (resolved === null) {
      throw new Error(`go to source ${entry.label}: source missing`);
    }
    await requireVisibleSheet(context, resolved, entry.label);
    if (resolved.kind === "chart") {
      resolved.chart.worksheet.activate();
      resolved.chart.activate();
    } else {
      // Excel refuses to select on a sheet that is not the active one.
      resolved.range.worksheet.activate();
      resolved.range.select();
    }
    await context.sync();
  });
}

// The relay copy is revoked first. A revoke that fails throws before anything
// local moves, so the entry - and the token the revoke needs - is still there
// to try again with; the other order leaves the relay serving a picture the
// workbook no longer knows how to withdraw.
export async function removeLink(id: string, relay: RelayApi): Promise<void> {
  await exclusive("remove", () =>
    Excel.run(async (context) => {
      const registry = await readRegistry(context);
      const entry = entryOf(registry, id, "remove");
      await forget(entry, relay);
      await releaseAnchor(context, entry);
      registry.links = registry.links.filter((link) => link.id !== id);
      writeRegistry(context, registry);
      await context.sync();
    }),
  );
}

// The pane's chart list follows the active sheet. Older hosts (below ExcelApi
// 1.7) have no worksheet activation event; the tab-open refresh covers them.
export function watchActiveSheet(handler: () => Promise<void>): void {
  if (!hostSupports("1.7")) return;
  Excel.run(async (context) => {
    context.workbook.worksheets.onActivated.add(async () => {
      await handler();
    });
    await context.sync();
  }).catch(() => undefined);
}
