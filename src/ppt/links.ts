// Link orchestration for the PowerPoint pane: what the deck holds crossed with
// what the relay knows, and the update, inbox and insert flows built on it.
// No Office.js - the host adapter is injectable, so every rule here is tested
// against a fake deck and a fake relay.

import { deriveLinkKeys, open } from "../link/crypto";
import {
  decodeInboxItem,
  decodePayload,
  type InboxItem,
  type Payload,
} from "../link/model";
import { RelayError, type RelayApi } from "../link/relay";
import {
  deriveStatus,
  sourceChanged,
  type LinkStatus,
  type RelayStatus,
} from "../link/status";
import type { Workspace } from "../link/workspace";
import * as realHost from "./host";
import type { FoundLink } from "./host";

export type PptHost = Pick<
  typeof realHost,
  "scanLinks" | "insertLink" | "refreshLink" | "breakLink" | "goToSlide"
>;

export interface LinkRow {
  found: FoundLink;
  status: LinkStatus;
  relayRev: number | null;
  pushedAt: number | null;
}

export interface UpdateSummary {
  updated: number;
  current: number;
  missing: number;
  wrongKey: number;
  failed: number;
  sourceChanges: string[];
}

// One auth per distinct token, derived once: a deck can hold the same link on
// twenty slides, and HKDF is not free.
async function authsFor(found: FoundLink[]): Promise<Map<string, string>> {
  const auths = new Map<string, string>();
  for (const link of found) {
    if (!auths.has(link.token)) {
      auths.set(link.token, (await deriveLinkKeys(link.token)).auth);
    }
  }
  return auths;
}

// The relay answers per (id, auth) and in request order: copies of a shape
// share both and are queried once; a re-keyed copy gets its own query.
export async function listLinks(
  relay: RelayApi,
  host: PptHost = realHost,
): Promise<LinkRow[]> {
  const found = await host.scanLinks();
  if (found.length === 0) return [];
  const auths = await authsFor(found);
  const pairKey = (link: FoundLink): string =>
    `${link.tag.id}/${auths.get(link.token)!}`;
  const pairs = [
    ...new Map<string, FoundLink>(
      found.map((link) => [pairKey(link), link]),
    ).values(),
  ];
  const statuses = await relay.status(
    pairs.map((link) => ({ id: link.tag.id, auth: auths.get(link.token)! })),
  );
  const byPair = new Map<string, RelayStatus | undefined>(
    pairs.map((link, index) => [pairKey(link), statuses[index]]),
  );
  return found.map((link) => {
    const status = byPair.get(pairKey(link));
    return {
      found: link,
      status: deriveStatus(link.tag.rev, status),
      relayRev: status?.rev ?? null,
      pushedAt: status?.pushedAt ?? null,
    };
  });
}

async function fetchPayload(
  found: FoundLink,
  relay: RelayApi,
): Promise<{ rev: number; payload: Payload } | "unchanged"> {
  const keys = await deriveLinkKeys(found.token);
  const result = await relay.getLink(found.tag.id, keys.auth, found.tag.rev);
  if (result === "unchanged") return "unchanged";
  return {
    rev: result.rev,
    payload: decodePayload(await open(keys.enc, found.tag.id, result.blob)),
  };
}

// A row the deck already agrees with is never fetched, and one row's failure
// never stops the rest: the summary is what the pane reports afterwards.
export async function updateLinks(
  rows: LinkRow[],
  relay: RelayApi,
  host: PptHost = realHost,
): Promise<UpdateSummary> {
  const summary: UpdateSummary = {
    updated: 0,
    current: 0,
    missing: 0,
    wrongKey: 0,
    failed: 0,
    sourceChanges: [],
  };
  for (const row of rows) {
    if (row.status !== "updateAvailable") {
      countSkipped(summary, row.status);
      continue;
    }
    try {
      const fetched = await fetchPayload(row.found, relay);
      if (fetched === "unchanged") {
        summary.current += 1;
        continue;
      }
      noteSourceChange(summary, row.found, fetched.payload);
      await host.refreshLink(row.found, fetched.payload, fetched.rev);
      summary.updated += 1;
    } catch (error) {
      countFailure(summary, error);
    }
  }
  return summary;
}

function countSkipped(summary: UpdateSummary, status: LinkStatus): void {
  if (status === "current") summary.current += 1;
  else if (status === "missing") summary.missing += 1;
  else if (status === "wrongKey") summary.wrongKey += 1;
}

// A link whose workbook changed still refreshes: the user is told which one,
// because a renamed file and the wrong file look the same from here.
function noteSourceChange(
  summary: UpdateSummary,
  found: FoundLink,
  payload: Payload,
): void {
  if (!sourceChanged(found.tag.src, payload.src)) return;
  summary.sourceChanges.push(
    `${found.tag.src.workbook} -> ${payload.src.workbook}`,
  );
}

function countFailure(summary: UpdateSummary, error: unknown): void {
  if (error instanceof RelayError && error.kind === "missing") {
    summary.missing += 1;
  } else if (error instanceof RelayError && error.kind === "auth") {
    summary.wrongKey += 1;
  } else {
    summary.failed += 1;
  }
}

export function summarize(summary: UpdateSummary): string {
  const parts = [
    [summary.updated, "updated"],
    [summary.current, "up to date"],
    [summary.missing, "missing"],
    [summary.wrongKey, "wrong key"],
    [summary.failed, "failed"],
  ] as const;
  const text = parts
    .filter(([count]) => count > 0)
    .map(([count, label]) => `${String(count)} ${label}`)
    .join(", ");
  return text || "No links found";
}

export async function listInbox(
  ws: Workspace,
  relay: RelayApi,
): Promise<InboxItem[]> {
  const items: InboxItem[] = [];
  for (const row of await relay.listInbox(ws.id, ws.auth)) {
    try {
      items.push(decodeInboxItem(await open(ws.enc, ws.id, row.blob)));
    } catch {
      // Sealed with another workspace key, or corrupt: not ours to show.
    }
  }
  return items;
}

export async function insertFromInbox(
  item: InboxItem,
  ws: Workspace,
  relay: RelayApi,
  host: PptHost = realHost,
): Promise<void> {
  const keys = await deriveLinkKeys(item.token);
  const result = await relay.getLink(item.id, keys.auth);
  if (result === "unchanged") {
    throw new Error(`insert ${item.label}: the relay returned no picture.`);
  }
  const payload = decodePayload(await open(keys.enc, item.id, result.blob));
  await host.insertLink(item, payload, result.rev);
  await relay.deleteInbox(ws.id, ws.auth, item.id);
}
