// Waits for a just-added PowerPoint shape to become visible to
// shapes.getItem(id), not just the id its own add already answered - the
// web can hold it out of both for a stretch (rig 27.09). Invariant: past
// the poll cap, a plain pane sentence, never a raw office.js string.

import { withSyncDeadline } from "./chart-draw";

// PowerPoint for the web registered a table 212-224 ms (3 polls) after the
// add's own sync answered its id; 30 tries at 100 ms apart covers that with
// margin - about 3 s worst case - before giving up.
export const REGISTER_POLL_MS = 100;
export const REGISTER_POLL_TRIES = 30;

// A real wait by default; a test swaps this for an instant resolve
// (setRegisterPollWait) so a poll that runs to the cap costs no real time.
let wait: (ms: number) => Promise<void> = (ms) =>
  new Promise((resolve) => setTimeout(resolve, ms));

export function setRegisterPollWait(fn: (ms: number) => Promise<void>): void {
  wait = fn;
}

// Exported so chart-cleanup.ts's own catch-up poll of still-missing ids waits
// the same way, without a second copy of the swap.
export function registerPollWait(ms: number): Promise<void> {
  return wait(ms);
}

// Polls shapes.getItemOrNullObject(id) until it answers present, at most
// REGISTER_POLL_TRIES times, REGISTER_POLL_MS apart. `what` should read like
// the add it follows: it names the round trip in the sentence a host that
// never answers at all gets, from withSyncDeadline.
export async function awaitShapeRegistered(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.ShapeCollection,
  id: string,
  what: string,
): Promise<void> {
  for (let tries = 0; tries < REGISTER_POLL_TRIES; tries += 1) {
    const probe = shapes.getItemOrNullObject(id);
    probe.load("isNullObject");
    await withSyncDeadline(context.sync(), what);
    if (!probe.isNullObject) return;
    if (tries < REGISTER_POLL_TRIES - 1) {
      await registerPollWait(REGISTER_POLL_MS);
    }
  }
  throw new Error("PowerPoint did not finish creating the table. Try again.");
}
