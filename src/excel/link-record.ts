// The relay round trip for one link: seal a rendered source into a payload
// and push it, seal an inbox note PowerPoint can find, and run both as one
// publish with rollback on failure. Anchors, registry and rendering stay in
// link-anchors.ts; this file only ever talks to the relay.

import { deriveLinkKeys, seal, sha256Hex } from "../link/crypto";
import {
  encodeInboxItem,
  encodePayload,
  type InboxItem,
  type Payload,
  type Registry,
  type RegistryEntry,
  type Source,
} from "../link/model";
import { base64ToBytes, pngSize } from "../link/png";
import type { RelayApi } from "../link/relay";
import type { Workspace } from "../link/workspace";
import { staged, writeRegistry } from "./link-anchors";
import type { Render } from "./link-render";

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

export async function pushPayload(
  entry: RegistryEntry,
  src: Source,
  render: Render,
  relay: RelayApi,
): Promise<number> {
  const payload = await payloadOf(src, render);
  const keys = await deriveLinkKeys(entry.token);
  const blob = await seal(keys.enc, entry.id, encodePayload(payload));
  return (await relay.putLink(entry.id, keys.auth, blob)).rev;
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
    };
  }
  if (render.kind === "text") {
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
  };
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
    throw staged(`export ${entry.label}`, error);
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
