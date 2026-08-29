// One queue for every flow that reads the PLSFIX_LINKS registry, changes it and
// writes the whole thing back: export, push, remove and the auto-push window.
// Each of those does network I/O between its read and its write - a getImage, a
// PUT to the relay, a POST to the inbox - so two of them overlapping means the
// second writes back a registry that predates the first. The link the export
// just recorded is then gone from the workbook while it still exists on the
// relay and in the deck, and no flow can reach the token that would revoke it.
//
// The Links tab's own busy state cannot close that window: the auto-push runs
// from a debounce timer, not from a click, so it never passes through the tab's
// interlock. This queue is the interlock both of them share.

// Resolved, never rejected: a flow that fails must not take the next one with
// it, so the chain always continues.
let queue: Promise<void> = Promise.resolve();
// What the queue is running, for a pane or a test that needs to say what an
// action is waiting behind; null while nothing holds it.
let holder: string | null = null;

export function linkQueueStage(): string | null {
  return holder;
}

// Runs work once every flow queued before it has finished, whatever the outcome
// of those. Never call it from inside work: the nested call would wait for a
// queue only its own caller can advance.
export function exclusive<T>(
  stage: string,
  work: () => Promise<T>,
): Promise<T> {
  const task = queue.then(async () => {
    holder = stage;
    try {
      return await work();
    } finally {
      holder = null;
    }
  });
  queue = task.then(
    () => undefined,
    () => undefined,
  );
  return task;
}
