// The relay round trip for one link: seal a rendered source into a payload
// and push it, seal an inbox note PowerPoint can find, and run both as one
// publish with rollback on failure. Anchors, registry and rendering stay in
// link-anchors.ts; this file only ever talks to the relay.

import { deriveLinkKeys, seal, sha256Hex } from "../link/crypto";
import { isLocalRev, localFloor } from "../link/local";
import {
  encodeInboxItem,
  encodePayload,
  TEXT_MAX_CHARS,
  TEXT_TOO_LONG,
  type InboxItem,
  type Payload,
  type Registry,
  type RegistryEntry,
  type Source,
} from "../link/model";
import { base64ToBytes, pngSize } from "../link/png";
import { isRelayError, type RelayApi } from "../link/relay";
import { relayReason } from "../link/relay-reason";
import type { Workspace } from "../link/workspace";
import { staged, writeRegistry } from "./link-anchors";
import { headerRow } from "./link-table";
import type { Render } from "./link-render";

// The one relay refusal a modeller can act on: the sealed export is past the
// relay's body limit (413), so a smaller range is the way out. Worded for
// this side on purpose - the pane is already in Excel, so it never says
// "from Excel" the way relay-reason.ts's shared TOO_LARGE does for
// PowerPoint's own toasts (src/ppt/links.audit.test.ts).
const TOO_BIG_TO_SEND =
  "That export is too big to send. Export a smaller range.";

// The one mapping both a push (pushOne, src/excel/links.ts) and an export's
// own first push (publish() below) need, so a 413/429/507/5xx never reaches
// a toast as "relay PUT /api/links/x: 507". undefined means the relay's own
// message should travel on unchanged.
export function relayFailureReason(error: unknown): string | undefined {
  return isRelayError(error) && error.kind === "tooLarge"
    ? TOO_BIG_TO_SEND
    : relayReason(error);
}

export interface NewLink {
  entry: RegistryEntry;
  src: Source;
  render: Render;
  // The registry as it was read before the anchor was bound: what publish
  // appends to, and what a rollback puts back.
  registry: Registry;
  // Undoes the anchor: deletes the hidden name, or gives the chart its own
  // name back. Queued only - the caller's next sync commits it.
  release: () => void;
}

// Which links this running pane has already pushed at least once: cleared
// by nothing but a fresh module load, which is exactly what a reload after
// an upgrade, a reopened workbook's pane starting cold, or a Save As copy's
// own separate pane amount to. localFloor (src/link/local.ts) trusts
// entry.rev/entry.localRev only once a link's floor has already passed
// through this session; the first time, it does not, however local-looking
// the registry's own number is.
const trustedThisSession = new Set<string>();

// The registry's own `rev` is not enough to count on from in local mode: a
// relay push in between resets it to the relay's own small counter, and
// counting on from that alone would have the very next local push land back
// on LOCAL_REV_BASE + 1 - a rev this link may already have used under
// different content, which src/link/local-store.ts's mergeRevision then
// reads as "nothing new" and keeps the stale blob. `entry.localRev` is the
// high-water mark that survives a relay excursion, so the floor a local push
// counts on from is whichever of the two is higher; a relay push (`localRev`
// stays untouched, per isLocalRev's own guard below) never gets to lower it.
// localFloor then covers what a surviving localRev cannot within one
// session: a registry whose own memory of that high-water mark is gone
// (predates the field, an unsaved reopen, a Save As copy), where the floor
// above comes back small, or local-looking but stale, with nothing in this
// fresh session yet to tell the two apart.
export async function pushPayload(
  entry: RegistryEntry,
  src: Source,
  render: Render,
  relay: RelayApi,
): Promise<number> {
  const payload = await payloadOf(src, render);
  const keys = await deriveLinkKeys(entry.token);
  const blob = await seal(keys.enc, entry.id, encodePayload(payload));
  const trusted = trustedThisSession.has(entry.id);
  trustedThisSession.add(entry.id);
  const floor = localFloor(Math.max(entry.rev, entry.localRev ?? 0), trusted);
  const rev = (await relay.putLink(entry.id, keys.auth, blob, floor)).rev;
  if (isLocalRev(rev)) entry.localRev = rev;
  return rev;
}

// The hash is what tells one push from the next: the picture itself, the cells
// of a table or the text of one cell - never the envelope around them, which
// carries the clock.
async function payloadOf(src: Source, render: Render): Promise<Payload> {
  const pushedAt = new Date().toISOString();
  if (render.kind === "table") {
    const { rows, cols, cells, widths } = render;
    return {
      v: 1,
      kind: "table",
      rows,
      cols,
      cells,
      widths,
      src,
      pushedAt,
      hash: await sha256Hex(JSON.stringify(cells)),
      // The plain fallback's cells never carry `b`, so a source Excel
      // refused the formats for never reads as a header either.
      ...(headerRow(cells) ? { h: true as const } : {}),
    };
  }
  if (render.kind === "text") {
    // Checked again here, not only at export time (link-export.ts's
    // requireTextCell): the anchor is a fixed named range, but the cell's
    // own text is not, so a later push can read a cell that grew past the
    // cap with no new export in between. The table branch above re-checks
    // its own cap the same way, inside renderTable.
    if (render.text.length > TEXT_MAX_CHARS) throw new Error(TEXT_TOO_LONG);
    return {
      v: 1,
      kind: "text",
      text: render.text,
      src,
      pushedAt,
      hash: await sha256Hex(render.text),
    };
  }
  const size = pngSize(base64ToBytes(render.png));
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: size.width,
    height: size.height,
    png: render.png,
    src,
    pushedAt,
    hash: await sha256Hex(render.png),
    ...(render.chart ? { chart: render.chart } : {}),
    ...(render.chartIssue === undefined
      ? {}
      : { chartIssue: render.chartIssue }),
  };
}

// The note PowerPoint picks up: it carries the token, which is why it is sealed
// with the workspace key the two panes were paired with. Exported for a
// Copy selected / Copy all in local mode, which re-announces every pushed
// link so a deck that does not hold it yet can still insert it.
export async function announce(
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
    ...(entry.project ? { project: entry.project } : {}),
  };
  const blob = await seal(ws.enc, ws.id, encodeInboxItem(item));
  await relay.postInbox(ws.id, ws.auth, entry.id, blob);
}

// The anchor is already bound when this runs - it has to be, so the picture and
// the name describe the same object - so every failure from here on has to put
// the workbook back rather than leave an anchor nothing points at.
export async function publish(
  context: Excel.RequestContext,
  link: NewLink,
  ws: Workspace,
  relay: RelayApi,
): Promise<void> {
  const { entry, src, registry } = link;
  let recorded = false;
  try {
    entry.rev = await pushPayload(entry, src, link.render, relay);
    entry.lastPushedAt = new Date().toISOString();
    // Recorded before the inbox note goes out, so PowerPoint is never told
    // about a link this workbook has no record of.
    writeRegistry(context, { ...registry, links: [...registry.links, entry] });
    recorded = true;
    await context.sync();
    await announce(entry, src, ws, relay);
  } catch (error) {
    await rollback(context, link, recorded);
    const stage = `export ${entry.label}`;
    const worded = relayFailureReason(error);
    throw staged(stage, worded ? new Error(worded) : error);
  }
}

// The setting is only rewritten if this flow had already written it: a failed
// export leaves a workbook that never had a registry exactly as it was. Best
// effort by design - if the workbook will not take the undo, the export error
// the caller is about to see is the one worth reporting.
async function rollback(
  context: Excel.RequestContext,
  link: NewLink,
  recorded: boolean,
): Promise<void> {
  try {
    link.release();
    if (recorded) writeRegistry(context, link.registry);
    await context.sync();
  } catch {
    return;
  }
}
