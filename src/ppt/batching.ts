// How much picture one repaint round trip is allowed to carry. "Update all" on
// a real pitch book can earn tens of megabytes of PNG, and every payload in one
// PowerPoint.run is held twice - once as the base64 the relay answered with,
// once inside the request the host is handed - so the repaints are cut into
// budget-sized runs instead of one run of any size. Greedy and in order, so the
// rows repaint in the order they were fetched, and a picture past the budget on
// its own still travels, alone. Pure: keys and byte counts, no payloads and no
// Office.js, so the plan is tested without a deck.

// Eight MiB of payload per PowerPoint.run: twice the relay's 4 MiB response
// cap, so a batch fetch that came back full still repaints in one round trip,
// and a deck of any size costs memory in that unit rather than in decks.
export const REPAINT_BUDGET_BYTES = 8 * 1024 * 1024;

// One thing to batch: whatever the caller finds it again by, and what it
// weighs.
export interface BatchItem {
  key: string;
  bytes: number;
}

// The keys in batches, in order, each batch inside the budget unless it holds a
// single item that is over it on its own. Every key comes back exactly once, so
// the caller maps straight back onto its own list.
export function planBatches(
  items: BatchItem[],
  budgetBytes: number,
): string[][] {
  const batches: string[][] = [];
  let batch: string[] = [];
  let used = 0;
  for (const item of items) {
    // Only a batch that holds something is flushed: an item bigger than the
    // whole budget would otherwise flush empty batches forever, and it has to
    // be painted somewhere.
    if (batch.length > 0 && used + item.bytes > budgetBytes) {
      batches.push(batch);
      batch = [];
      used = 0;
    }
    batch.push(item.key);
    used += item.bytes;
  }
  if (batch.length > 0) batches.push(batch);
  return batches;
}
