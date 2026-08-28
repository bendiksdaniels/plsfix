// Link orchestration for the PowerPoint pane: what the deck holds crossed with
// what the relay knows, and the update, inbox and insert flows built on it.
// No Office.js - the host adapter is injectable, so every rule here is tested
// against a fake deck and a fake relay.

import { deriveLinkKeys, open } from "../link/crypto";
import {
  decodeInboxItem,
  decodePayload,
  sourceLabel,
  type InboxItem,
  type Payload,
} from "../link/model";
import {
  isRelayError,
  RelayError,
  type RelayApi,
  type StatusQuery,
} from "../link/relay";
import {
  deriveStatus,
  sourceChanged,
  type LinkStatus,
  type RelayStatus,
} from "../link/status";
import type { Workspace } from "../link/workspace";
import * as realHost from "./host";
import type { FoundLink, RefreshRequest } from "./host";

// refreshLinks is optional: a stub host is a handful of functions, and without
// it every row is simply refreshed on its own - what a host below
// PowerPointApi 1.8 does anyway.
export type PptHost = Pick<
  typeof realHost,
  "scanLinks" | "insertLink" | "refreshLink" | "breakLink" | "goToSlide"
> & { refreshLinks?: typeof realHost.refreshLinks };

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
  failures: string[];
}

interface Keyed {
  link: FoundLink;
  auth: string | null;
}

// One auth per distinct token, derived once: a deck can hold the same link on
// twenty slides, and HKDF is not free. A token that will not derive is a broken
// key on that one shape, not a broken pane - it gets a null auth, reports
// wrongKey below and never reaches the relay, and its neighbours carry on.
async function keyLinks(found: FoundLink[]): Promise<Keyed[]> {
  const auths = new Map<string, string | null>();
  for (const link of found) {
    if (auths.has(link.token)) continue;
    try {
      auths.set(link.token, (await deriveLinkKeys(link.token)).auth);
    } catch {
      auths.set(link.token, null);
    }
  }
  return found.map((link) => ({ link, auth: auths.get(link.token) ?? null }));
}

function pairKey(id: string, auth: string): string {
  return `${id}/${auth}`;
}

// The relay answers per (id, auth) and in request order: copies of a shape
// share both and are queried once; a re-keyed copy gets its own query. A reply
// that does not line up row for row cannot be read positionally at all.
async function statusByPair(
  keyed: Keyed[],
  relay: RelayApi,
): Promise<Map<string, RelayStatus | undefined>> {
  const pairs = new Map<string, StatusQuery>();
  for (const { link, auth } of keyed) {
    if (auth !== null)
      pairs.set(pairKey(link.tag.id, auth), { id: link.tag.id, auth });
  }
  if (pairs.size === 0) return new Map();
  const keys = [...pairs.keys()];
  const statuses = await relay.status([...pairs.values()]);
  if (statuses.length !== keys.length) {
    throw new RelayError(
      "server",
      `relay status: expected ${String(keys.length)} rows, got ${String(statuses.length)}`,
    );
  }
  return new Map<string, RelayStatus | undefined>(
    keys.map((key, index) => [key, statuses[index]]),
  );
}

export async function listLinks(
  relay: RelayApi,
  host: PptHost = realHost,
): Promise<LinkRow[]> {
  const found = await host.scanLinks();
  if (found.length === 0) return [];
  const keyed = await keyLinks(found);
  const byPair = await statusByPair(keyed, relay);
  return keyed.map(({ link, auth }) => {
    const status =
      auth === null ? undefined : byPair.get(pairKey(link.tag.id, auth));
    return {
      found: link,
      status: auth === null ? "wrongKey" : deriveStatus(link.tag.rev, status),
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
// never stops the rest: the summary is what the pane reports afterwards. The
// fetches run row by row, the repaints they earn all travel together.
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
    failures: [],
  };
  const batch: RefreshRequest[] = [];
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
      batch.push({
        found: row.found,
        payload: fetched.payload,
        rev: fetched.rev,
      });
    } catch (error) {
      countFailure(summary, row.found, error);
    }
  }
  await applyRefreshes(summary, batch, host);
  return summary;
}

// Every picture the relay had something new for, repainted in one round trip.
// The batch is a speed-up, never a new failure mode: a host that refuses one
// shape rejects the whole run, so the rows go through one at a time and only
// the bad one is counted as a failure. Replaying a row the batch had already
// applied costs nothing - a repaint writes the same picture, tag and height.
async function applyRefreshes(
  summary: UpdateSummary,
  batch: RefreshRequest[],
  host: PptHost,
): Promise<void> {
  if (batch.length === 0) return;
  if (host.refreshLinks) {
    try {
      if (await host.refreshLinks(batch)) {
        summary.updated += batch.length;
        return;
      }
    } catch {
      // One shape in the batch; the rows below name it.
    }
  }
  for (const entry of batch) {
    try {
      await host.refreshLink(entry.found, entry.payload, entry.rev);
      summary.updated += 1;
    } catch (error) {
      countFailure(summary, entry.found, error);
    }
  }
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

// A relay answer the pane already has a column for is only counted; anything
// else keeps its message, because "3 failed" with no reason is unactionable.
function countFailure(
  summary: UpdateSummary,
  found: FoundLink,
  error: unknown,
): void {
  const kind = isRelayError(error) ? error.kind : null;
  if (kind === "missing") {
    summary.missing += 1;
  } else if (kind === "auth") {
    summary.wrongKey += 1;
  } else {
    summary.failed += 1;
    summary.failures.push(failureLine(found, error));
  }
}

// Every failure names its link. The host already stages its own errors as
// "refresh <label>: ...", so the label is not stuttered back onto those.
function failureLine(found: FoundLink, error: unknown): string {
  const label = sourceLabel(found.tag.src, found.tag.kind);
  const message = error instanceof Error ? error.message : String(error);
  return message.includes(label) ? message : `${label}: ${message}`;
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
