// Link orchestration for the PowerPoint pane: what the deck holds crossed with
// what the relay knows, and the update, inbox and insert flows built on it.
// No Office.js - the host adapter is injectable, so every rule here is tested
// against a fake deck and a fake relay.

import { deriveLinkKeys, open } from "../link/crypto";
import {
  decodeInboxItem,
  decodePayload,
  payloadBytes,
  sourceLabel,
  type InboxItem,
  type Payload,
} from "../link/model";
import {
  isRelayError,
  MAX_STATUS_ITEMS,
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
import { chunk, planBatches, REPAINT_BUDGET_BYTES } from "./batching";
import { fetchUpdates } from "./fetch";
import * as realHost from "./host";
import type { FoundLink, InsertResult, RefreshRequest } from "./host";

// refreshLinks is optional: a stub host is a handful of functions, and without
// it every row is simply refreshed on its own - what a host below
// PowerPointApi 1.8 does anyway.
export type PptHost = Pick<
  typeof realHost,
  | "scanLinks"
  | "insertLink"
  | "refreshLink"
  | "retagLink"
  | "breakLink"
  | "goToSlide"
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
// that does not line up row for row cannot be read positionally at all, so the
// count is checked per request. A deck can hold more distinct pairs than one
// request may carry - the relay refuses a longer batch with a 400, which used
// to leave a big pitch book unable to list a single link - so the pairs are
// asked about MAX_STATUS_ITEMS at a time and each answer is matched against
// the slice that earned it.
async function statusByPair(
  keyed: Keyed[],
  relay: RelayApi,
): Promise<Map<string, RelayStatus | undefined>> {
  const pairs = new Map<string, StatusQuery>();
  for (const { link, auth } of keyed) {
    if (auth !== null)
      pairs.set(pairKey(link.tag.id, auth), { id: link.tag.id, auth });
  }
  const byPair = new Map<string, RelayStatus | undefined>();
  for (const part of chunk([...pairs], MAX_STATUS_ITEMS)) {
    const statuses = await relay.status(part.map(([, query]) => query));
    if (statuses.length !== part.length) {
      throw new RelayError(
        "server",
        `relay status: expected ${String(part.length)} rows, got ${String(statuses.length)}`,
      );
    }
    part.forEach(([key], index) => byPair.set(key, statuses[index]));
  }
  return byPair;
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

// A row the deck already agrees with is never fetched, and one row's failure
// never stops the rest: the summary is what the pane reports afterwards. The
// fetches travel in one batch (`fetch.ts`) and the repaints they earn in
// byte-budgeted batches below, so a deck of any size costs a handful of round
// trips, not two per link.
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
  const wanted = rows.filter((row) => {
    if (row.status === "updateAvailable") return true;
    countSkipped(summary, row.status);
    return false;
  });
  const fetched = await fetchUpdates(wanted, relay);
  summary.current += fetched.current;
  summary.missing += fetched.missing;
  summary.wrongKey += fetched.wrongKey;
  for (const { found, error } of fetched.failures) {
    countFailure(summary, found, error);
  }
  for (const entry of fetched.batch) {
    noteSourceChange(summary, entry.found, entry.payload);
  }
  summary.updated += await applyBatch(fetched.batch, host, (found, error) => {
    countFailure(summary, found, error);
  });
  return summary;
}

// Every picture a run has something to paint, and the count of the ones that
// landed. The repaints travel in as few round trips as their bytes allow: one
// host batch per REPAINT_BUDGET_BYTES of payload, because a deck can earn more
// picture than a single host request has any business carrying. Shared with
// revert.ts, which paints an older revision through exactly the same path.
export async function applyBatch(
  batch: RefreshRequest[],
  host: PptHost,
  onFailure: (found: FoundLink, error: unknown) => void,
): Promise<number> {
  let painted = 0;
  for (const part of splitByBytes(batch)) {
    painted += await paintBatch(part, host, onFailure);
  }
  return painted;
}

// The payload list cut to the repaint budget. The key is the row's position, so
// two shapes carrying the same picture still map back to their own entries, and
// the order the rows were fetched in is the order they repaint in.
function splitByBytes(batch: RefreshRequest[]): RefreshRequest[][] {
  const items = batch.map((entry, index) => ({
    key: String(index),
    // What the host request carries is what the budget counts: a picture's
    // base64, about four bytes for every three of picture, or a table's cells.
    bytes: payloadBytes(entry.payload),
  }));
  return planBatches(items, REPAINT_BUDGET_BYTES).map((keys) =>
    keys.map((key) => batch[Number(key)]!),
  );
}

// One batch in one round trip. The batch is a speed-up, never a new failure
// mode: a host that refuses one shape rejects the whole run it sits in, so
// those rows go through one at a time, only the bad one is reported, and the
// other batches never notice. Replaying a row this run had already applied
// costs nothing - a repaint writes the same picture, tag and height.
async function paintBatch(
  batch: RefreshRequest[],
  host: PptHost,
  onFailure: (found: FoundLink, error: unknown) => void,
): Promise<number> {
  if (host.refreshLinks) {
    try {
      if (await host.refreshLinks(batch)) return batch.length;
    } catch {
      // One shape in the batch; the rows below name it.
    }
  }
  let painted = 0;
  for (const entry of batch) {
    try {
      await host.refreshLink(entry.found, entry.payload, entry.rev);
      painted += 1;
    } catch (error) {
      onFailure(entry.found, error);
    }
  }
  return painted;
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
export function failureLine(found: FoundLink, error: unknown): string {
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
  const items: { item: InboxItem; createdAt: number }[] = [];
  for (const row of await relay.listInbox(ws.id, ws.auth)) {
    try {
      items.push({
        item: decodeInboxItem(await open(ws.enc, ws.id, row.blob)),
        createdAt: row.createdAt,
      });
    } catch {
      // Sealed with another workspace key, or corrupt: not ours to show.
    }
  }
  // A relay is free to return rows in storage order. The pane's one-click
  // paste must mean newest export, not whichever row happened to arrive first.
  return items
    .sort((left, right) => right.createdAt - left.createdAt)
    .map(({ item }) => item);
}

export function latestInboxItem(items: InboxItem[]): InboxItem | null {
  return items[0] ?? null;
}

// What the pane adds to "Inserted <label>." after a host that had something to
// say: where the object had to land, then why a chart came as a picture.
export function insertNote(placed: InsertResult): string {
  const parts = [
    placed.overlapping ? realHost.OVERLAP_NOTE : null,
    placed.note ?? null,
  ].filter((part): part is string => part !== null);
  return parts.length === 0 ? "" : ` ${parts.join(". ")}`;
}

export async function insertFromInbox(
  item: InboxItem,
  ws: Workspace,
  relay: RelayApi,
  host: PptHost = realHost,
): Promise<InsertResult> {
  const keys = await deriveLinkKeys(item.token);
  const result = await relay.getLink(item.id, keys.auth);
  if (result === "unchanged") {
    throw new Error(`insert ${item.label}: the relay returned no picture.`);
  }
  const payload = decodePayload(await open(keys.enc, item.id, result.blob));
  const placed = await host.insertLink(item, payload, result.rev);
  await relay.deleteInbox(ws.id, ws.auth, item.id);
  return placed;
}
