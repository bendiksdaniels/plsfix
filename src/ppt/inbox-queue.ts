// The Inbox as this pane sees it: the relay's rows in the order a modeller
// means by "latest", minus the exports this pane has already pasted. Pure -
// no relay, no Office.js; links.ts does the fetching and the inserting.
// Invariant: the relay drops a row as its insert lands, but a delete it
// refuses must not have the next one-click paste hand back the same export
// and land a second copy, so an id pasted here is never offered again.

import type { InboxItem } from "../link/model";

const pasted = new Set<string>();

export function rememberPasted(id: string): void {
  pasted.add(id);
}

// The relay orders rows newest first with ties broken by arrival, and a
// self-hosted one is free to answer in storage order: the pane's one-click
// paste must mean newest export, so the sort is repeated here; it is stable,
// so equal createdAt values keep the server's order.
export function queueOrder(
  rows: { item: InboxItem; createdAt: number }[],
): InboxItem[] {
  return rows
    .sort((left, right) => right.createdAt - left.createdAt)
    .map(({ item }) => item)
    .filter((item) => !pasted.has(item.id));
}

export function latestInboxItem(items: InboxItem[]): InboxItem | null {
  return items[0] ?? null;
}
