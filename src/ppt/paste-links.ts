// src/ppt/paste-links.ts
// Local mode's paste flow: a bundle pasted from Excel goes into the store,
// then repaints whatever of it is already on a slide. Owns pasteLinks and
// its one-line summary. Invariant: one deck scan per paste (the host is
// never asked twice); a link the deck does not hold yet only ever reaches
// the Inbox here, never the slide.

import type { Bundle } from "../link/bundle";
import type { LocalStore } from "../link/local-store";
import * as realHost from "./host";
import { listLinks, updateLinks, type PptHost } from "./links";

export interface PasteSummary {
  links: number; // links the bundle carried
  updated: number;
  current: number;
  waiting: number; // new inbox rows kept for links the deck does not hold
  failed: number;
  failures: string[];
  notes: string[];
}

export async function pasteLinks(
  bundle: Bundle,
  store: LocalStore,
  host: PptHost = realHost,
): Promise<PasteSummary> {
  const found = await host.scanLinks();
  const deckIds = new Set(found.map((link) => link.tag.id));
  const ingested = await store.ingest(bundle, deckIds);
  // One scan for the whole paste: listLinks would otherwise scan the deck a
  // second time to build the rows updateLinks repaints from.
  const once: PptHost = { ...host, scanLinks: async () => found };
  const rows = await listLinks(store, once);
  const pastedIds = new Set(ingested.linkIds);
  const subset = rows.filter((row) => pastedIds.has(row.found.tag.id));
  const summary = await updateLinks(subset, store, host);
  return {
    links: bundle.links.length,
    updated: summary.updated,
    current: summary.current,
    waiting: ingested.waiting,
    // missing/wrongKey/notPasted are not reachable for a row this same paste
    // just ingested (kept here only so a broken token on the shape still
    // counts as a failure rather than vanishing from the total).
    failed:
      summary.failed + summary.missing + summary.wrongKey + summary.notPasted,
    failures: summary.failures,
    notes: summary.notes,
  };
}

export function summarizePaste(summary: PasteSummary): string {
  const parts = [
    [summary.updated, "updated"],
    [summary.current, "up to date"],
    [summary.waiting, "waiting in the Inbox"],
  ] as const;
  const text = parts
    .filter(([count]) => count > 0)
    .map(([count, label]) => `${String(count)} ${label}`)
    .join(", ");
  const failedPart =
    summary.failed > 0
      ? `${text === "" ? "" : ", "}${String(summary.failed)} failed`
      : "";
  const body = `${text}${failedPart}`;
  if (summary.links === 0) return body === "" ? "Pasted." : `Pasted: ${body}.`;
  const noun = summary.links === 1 ? "link" : "links";
  const prefix = `Pasted ${String(summary.links)} ${noun}`;
  return body === "" ? `${prefix}.` : `${prefix}: ${body}.`;
}
