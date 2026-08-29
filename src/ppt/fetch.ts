// The fetch half of "Update all": every changed picture of a deck in one relay
// round trip instead of one GET per link. Rows are grouped per link and key (a
// link that sits on twenty slides is asked for once), chunked to the relay's
// item limit and opened here. The batch is a speed-up, never a new failure
// mode: whatever the relay defers past its response cap - and a whole batch it
// refused outright - falls back to the per-row GET this replaced. No Office.js,
// so the relay is injectable like everything else in this folder.
//
// Two ceilings bound an update-all, and they are deliberately not the same one:
//   - the wire, here: MAX_FETCH_ITEMS (200) queries per request, so a big deck
//     asks in several; and FETCH_BLOB_CAP (4 MiB) of blobs per answer, which
//     the relay enforces by calling the overflow "deferred" - those links are
//     then fetched one GET each below.
//   - the repaint, in `batching.ts`: REPAINT_BUDGET_BYTES (8 MiB) of payload
//     per PowerPoint.run, applied by `applyBatch` once everything here has been
//     opened. Twice the response cap on purpose, so a full batch answer still
//     repaints in a single round trip.
// Neither bound counts the other's bytes: the first keeps one HTTP response
// sane, the second keeps one host request out of the deck's memory.

import { deriveLinkKeys, open, type LinkKeys } from "../link/crypto";
import { decodePayload } from "../link/model";
import {
  MAX_FETCH_ITEMS,
  type FetchedLink,
  type FetchQuery,
  type FetchResult,
  type OmittedReason,
  type RelayApi,
} from "../link/relay";
import type { FoundLink, RefreshRequest } from "./host";
import type { LinkRow } from "./links";

// What a run of fetches earned: the repaints to apply, the rows that need
// none, and the ones that failed with a reason the caller turns into a line
// of the update summary.
export interface FetchOutcome {
  batch: RefreshRequest[];
  current: number;
  missing: number;
  wrongKey: number;
  failures: { found: FoundLink; error: unknown }[];
}

// One query per (link id, auth key): copies of a shape share both and are
// asked for once, and a re-keyed copy asks on its own. `knownRev` is the
// deck's revision only while every copy agrees on it; otherwise it is left
// out, so the relay answers with the picture rather than calling one of the
// copies up to date.
interface Group {
  id: string;
  keys: LinkKeys;
  knownRev: number | undefined;
  rows: LinkRow[];
}

function emptyOutcome(): FetchOutcome {
  return { batch: [], current: 0, missing: 0, wrongKey: 0, failures: [] };
}

// One derivation per distinct token: a deck can hold the same link on twenty
// slides, and HKDF is not free. A token that will not derive is a broken key
// on that one shape, not a broken pane.
async function keysByToken(
  rows: LinkRow[],
): Promise<Map<string, LinkKeys | null>> {
  const keys = new Map<string, LinkKeys | null>();
  for (const row of rows) {
    if (keys.has(row.found.token)) continue;
    try {
      keys.set(row.found.token, await deriveLinkKeys(row.found.token));
    } catch {
      keys.set(row.found.token, null);
    }
  }
  return keys;
}

async function groupRows(
  rows: LinkRow[],
  outcome: FetchOutcome,
): Promise<Group[]> {
  const keys = await keysByToken(rows);
  const groups = new Map<string, Group>();
  for (const row of rows) {
    const linkKeys = keys.get(row.found.token) ?? null;
    if (linkKeys === null) {
      outcome.wrongKey += 1;
      continue;
    }
    const id = row.found.tag.id;
    const group = groups.get(`${id}/${linkKeys.auth}`);
    if (group === undefined) {
      groups.set(`${id}/${linkKeys.auth}`, {
        id,
        keys: linkKeys,
        knownRev: row.found.tag.rev,
        rows: [row],
      });
      continue;
    }
    group.rows.push(row);
    if (group.knownRev !== row.found.tag.rev) group.knownRev = undefined;
  }
  return [...groups.values()];
}

// One request per 200 groups, and never two groups of the same link id in one
// request: the answer names items by id alone, so a second key on a shared id
// waits for its own GET instead of reading the first one's blob.
function chunk(groups: Group[]): { batches: Group[][]; deferred: Group[] } {
  const batches: Group[][] = [];
  const deferred: Group[] = [];
  const claimed = new Set<string>();
  let batch: Group[] = [];
  for (const group of groups) {
    if (claimed.has(group.id)) {
      deferred.push(group);
      continue;
    }
    claimed.add(group.id);
    batch.push(group);
    if (batch.length === MAX_FETCH_ITEMS) {
      batches.push(batch);
      batch = [];
    }
  }
  if (batch.length > 0) batches.push(batch);
  return { batches, deferred };
}

function query(group: Group): FetchQuery {
  const { id, keys, knownRev } = group;
  return knownRev === undefined
    ? { id, auth: keys.auth }
    : { id, auth: keys.auth, knownRev };
}

// One picture, opened once and repainted onto every shape that holds it. A
// blob this key cannot open is that link's failure, not the batch's.
async function openInto(
  group: Group,
  item: FetchedLink,
  outcome: FetchOutcome,
): Promise<void> {
  try {
    const payload = decodePayload(
      await open(group.keys.enc, group.id, item.blob),
    );
    for (const row of group.rows) {
      outcome.batch.push({ found: row.found, payload, rev: item.rev });
    }
  } catch (error) {
    for (const row of group.rows)
      outcome.failures.push({ found: row.found, error });
  }
}

// The relay's reasons, in the pane's counters. "deferred" is the one that is
// not an answer: those rows are fetched on their own below.
function countOmitted(
  group: Group,
  reason: OmittedReason,
  outcome: FetchOutcome,
  deferred: Group[],
): void {
  if (reason === "deferred") deferred.push(group);
  else if (reason === "unchanged") outcome.current += group.rows.length;
  else if (reason === "missing") outcome.missing += group.rows.length;
  else outcome.wrongKey += group.rows.length;
}

async function fetchBatch(
  batch: Group[],
  relay: RelayApi,
  outcome: FetchOutcome,
  deferred: Group[],
): Promise<void> {
  const waiting = new Map(batch.map((group) => [group.id, group]));
  let answer: FetchResult;
  try {
    answer = await relay.fetchLinks(batch.map(query));
  } catch {
    // A batch the relay refused says nothing about any one link - an older
    // relay does not know the route at all - so every row falls back to the
    // GET it would have made, which reports its own reason if the relay
    // really is unreachable.
    deferred.push(...batch);
    return;
  }
  for (const item of answer.items) {
    const group = waiting.get(item.id);
    if (group === undefined) continue;
    waiting.delete(item.id);
    await openInto(group, item, outcome);
  }
  for (const { id, reason } of answer.omitted) {
    const group = waiting.get(id);
    if (group === undefined) continue;
    waiting.delete(id);
    countOmitted(group, reason, outcome, deferred);
  }
  // A link the relay answered neither way is asked for on its own.
  deferred.push(...waiting.values());
}

// The path the batch replaced, kept for the rows it could not carry: one GET
// each, If-None-Match and all.
async function fetchRow(
  row: LinkRow,
  keys: LinkKeys,
  relay: RelayApi,
  outcome: FetchOutcome,
): Promise<void> {
  const { id, rev } = row.found.tag;
  try {
    const result = await relay.getLink(id, keys.auth, rev);
    if (result === "unchanged") {
      outcome.current += 1;
      return;
    }
    const payload = decodePayload(await open(keys.enc, id, result.blob));
    outcome.batch.push({ found: row.found, payload, rev: result.rev });
  } catch (error) {
    outcome.failures.push({ found: row.found, error });
  }
}

// Every row the relay has an update for, in as few round trips as it allows.
export async function fetchUpdates(
  rows: LinkRow[],
  relay: RelayApi,
): Promise<FetchOutcome> {
  const outcome = emptyOutcome();
  const { batches, deferred } = chunk(await groupRows(rows, outcome));
  for (const batch of batches)
    await fetchBatch(batch, relay, outcome, deferred);
  for (const group of deferred) {
    for (const row of group.rows)
      await fetchRow(row, group.keys, relay, outcome);
  }
  return outcome;
}
