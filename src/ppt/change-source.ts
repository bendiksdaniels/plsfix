// "Change source": point a tracked picture at a different export - usually the
// same table exported again from a newer workbook - keeping its slide, its
// position and its size. A link only re-points at an export of its own kind:
// a text box or a table cannot take a picture's payload in place, nor the
// other way round. The deck's tags are rewritten and the picture is
// repainted through the ordinary refresh path; nothing moves in Excel, and the
// link the shape used to hold stays on the relay untouched, so any other deck
// still tracking it carries on. No Office.js: the host and the relay are
// injectable, like everything else in this folder.

import { deriveLinkKeys, open } from "../link/crypto";
import {
  decodePayload,
  sourceLabel,
  type InboxItem,
  type LinkKind,
  type LinkTag,
  type Payload,
} from "../link/model";
import type { RelayApi } from "../link/relay";
import type { Workspace } from "../link/workspace";
import type { FoundLink } from "./host";
import * as realHost from "./host";
import { applyBatch, type LinkRow, type PptHost } from "./links";

const ONE_ROW = "Tick exactly one link to change its source.";
const NO_CANDIDATES =
  "Nothing waiting in the Inbox fits this link. Export the same kind of object again from Excel first.";
const NO_CHOICE = "Choose an export from the list.";

// Excel upper-cases refs and treats sheet names case-insensitively, and a
// workbook saved under another case is the same workbook - so every comparison
// here folds case, the way sourceChanged does.
function fold(value: string): string {
  return value.toLowerCase();
}

// How well an inbox item answers "this picture, from somewhere else": the
// label the row shows, then the anchor the export follows (or the same cells
// on the same sheet), then anything at all. Every item is offered - a
// re-exported table can carry a new label - but the likely one comes first.
function rank(row: LinkRow, item: InboxItem): number {
  const src = row.found.tag.src;
  if (fold(item.label) === fold(sourceLabel(src, row.found.tag.kind))) return 0;
  if (fold(item.src.anchor) === fold(src.anchor)) return 1;
  if (
    fold(item.src.sheet) === fold(src.sheet) &&
    fold(item.src.ref) === fold(src.ref)
  )
    return 1;
  return 2;
}

// What a shape can hold: a text box its text, a native table its cells, and a
// rectangle a picture - a range and a chart being the same rectangle, which is
// why they are the one pair that mixes.
function family(kind: LinkKind): LinkKind | "picture" {
  return kind === "text" || kind === "table" ? kind : "picture";
}

// No shape can take another family's payload in place. A table handed a
// picture kept its old cells and grew the new render behind them - a source
// changed, a row reading "up to date" and a slide still showing the previous
// workbook's numbers - and a rectangle handed a table's cells answers
// GeneralException, because getTable refuses it. Candidates stay on the row's
// side of both lines; kindWarning covers the range/chart pair that is left.
function sameFamily(row: LinkRow, item: InboxItem): boolean {
  return family(row.found.tag.kind) === family(item.kind);
}

// Pure, and stable inside a tier: the inbox's own order (newest first) decides
// between two equally good candidates.
export function candidatesFor(row: LinkRow, inbox: InboxItem[]): InboxItem[] {
  return inbox
    .map((item, index) => ({ item, rank: rank(row, item), index }))
    .filter((entry) => sameFamily(row, entry.item))
    .sort((left, right) => left.rank - right.rank || left.index - right.index)
    .map((entry) => entry.item);
}

// Re-pointing several pictures at one export is never what the user means, and
// with nothing ticked there is no picture to re-point.
export function requireOneRow(rows: LinkRow[]): LinkRow {
  const row = rows[0];
  if (row === undefined || rows.length > 1) throw new Error(ONE_ROW);
  return row;
}

export function requireCandidates(items: InboxItem[]): InboxItem[] {
  if (items.length === 0) throw new Error(NO_CANDIDATES);
  return items;
}

export function pickCandidate(items: InboxItem[], id: string): InboxItem {
  const item = items.find((candidate) => candidate.id === id);
  if (item === undefined) throw new Error(NO_CHOICE);
  return item;
}

// A range picture re-pointed at a chart (or the reverse) is allowed - the user
// may well mean it - but it is worth saying out loud, because the two render
// nothing alike and the row's label changes shape afterwards.
export function kindWarning(row: LinkRow, item: InboxItem): string | undefined {
  if (item.kind === row.found.tag.kind) return undefined;
  return `The new source is a ${item.kind}; this link tracked a ${row.found.tag.kind}.`;
}

function tagFor(item: InboxItem, payload: Payload, rev: number): LinkTag {
  return {
    v: 1,
    id: item.id,
    kind: item.kind,
    rev,
    src: payload.src,
    pushedAt: payload.pushedAt,
    ...(item.project ? { project: item.project } : {}),
  };
}

// The whole link the shape holds afterwards: the same shape, in the same place,
// carrying the new identity and the new key. Handing this to the repaint is
// what makes the reinsertion fallback below PowerPointApi 1.8 write the new
// tags rather than the old ones.
function retarget(found: FoundLink, tag: LinkTag, token: string): FoundLink {
  return { ...found, tag, token };
}

function message(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

// One repaint, through the shared batch path so a change of source paints
// exactly the way an update does. Zero painted is the failure the caller
// reports; the reason the host gave is kept.
async function repaint(
  found: FoundLink,
  payload: Payload,
  rev: number,
  host: PptHost,
  stage: string,
): Promise<void> {
  let failure: unknown = null;
  const painted = await applyBatch(
    [{ found, payload, rev }],
    host,
    (_found, error) => {
      failure = error;
    },
  );
  if (painted === 0) throw new Error(`${stage}: ${message(failure)}`);
}

// Tags first, picture second. Below PowerPointApi 1.8 the repaint reinserts the
// picture as a new shape and deletes the old one, so a retag afterwards would
// address a shape that is gone; and on 1.8 fill.setImage never writes PLSFIX_KEY
// at all, so this is the only thing that does. A repaint that fails puts the
// old tags back, because a shape claiming to be current while showing the
// previous source's picture is worse than a plain error.
export async function changeSource(
  row: LinkRow,
  item: InboxItem,
  ws: Workspace,
  relay: RelayApi,
  host: PptHost = realHost,
): Promise<string> {
  const before = row.found.tag.src.workbook;
  const stage = `change source ${sourceLabel(row.found.tag.src, row.found.tag.kind)}`;
  const keys = await deriveLinkKeys(item.token);
  // No knownRev: this deck has never held this link, so there is nothing the
  // relay could call unchanged.
  const result = await relay.getLink(item.id, keys.auth);
  if (result === "unchanged") {
    throw new Error(`${stage}: the relay returned no picture.`);
  }
  const payload = decodePayload(await open(keys.enc, item.id, result.blob));
  const tag = tagFor(item, payload, result.rev);
  const found = retarget(row.found, tag, item.token);
  await host.retagLink(found, tag, item.token);
  try {
    await repaint(found, payload, result.rev, host, stage);
  } catch (error) {
    await rollback(row, host);
    throw error;
  }
  await relay.deleteInbox(ws.id, ws.auth, item.id);
  return `Source changed: ${before} -> ${payload.src.workbook}`;
}

// Best effort by design: the repaint's failure is the one the user needs to
// read, so a rollback that fails too must not replace it. The row still shows
// the old source, which is what the deck is actually holding.
async function rollback(row: LinkRow, host: PptHost): Promise<void> {
  try {
    await host.retagLink(row.found, row.found.tag, row.found.token);
  } catch {
    // Reported by the next scan as whatever the shape now holds.
  }
}
