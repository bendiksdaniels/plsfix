// Boot-time TTL refresh: tells the relay which links this workbook still holds
// so their 30 days run from today, not from the last push. Reads the registry
// only; never writes it, never toasts, never blocks the pane.

import { deriveLinkKeys } from "../link/crypto";
import type { RelayApi, TouchQuery } from "../link/relay";
import { readRegistry } from "./link-anchors";

export async function touchWorkbookLinks(relay: RelayApi): Promise<number> {
  const registry = await Excel.run((context) => readRegistry(context));
  if (registry.links.length === 0) return 0;
  const items: TouchQuery[] = [];
  for (const entry of registry.links) {
    const keys = await deriveLinkKeys(entry.token);
    items.push({ id: entry.id, auth: keys.auth });
  }
  return relay.touchLinks(items);
}
