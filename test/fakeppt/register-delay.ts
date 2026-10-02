// Opt-in fake of PowerPoint for the web's registration window (rig 27.09,
// src/ppt/shape-ready.ts): a shape stays invisible to shapes.getItem(id)
// and the slide's item list for a stretch after its add. Off unless armed
// by helpers.delayRegistration; owns which ids are blocked, and for how long.

import type { FakePresentation } from "./model";

interface DelayState {
  // helpers.delayRegistration's one-shot arm, spent by the next add.
  armed: number | null;
  // Added while armed, waiting for its OWN add to be confirmed by a sync:
  // before that, nothing has re-anchored it yet, so it is not blocked at all.
  pending: Map<string, number>;
  // Confirmed and currently blocked: how many more successful syncs before
  // the host's own lookup catches up. Ages by one on every sync that lands,
  // whatever it was for - that passage of syncs IS the delay being modelled.
  blockedFor: Map<string, number>;
}

const decks = new WeakMap<FakePresentation, DelayState>();

function stateOf(deck: FakePresentation): DelayState {
  let state = decks.get(deck);
  if (!state) {
    state = { armed: null, pending: new Map(), blockedFor: new Map() };
    decks.set(deck, state);
  }
  return state;
}

// helpers.delayRegistration(syncs): the next shape objects.ts creates through
// addTable stays unregistered for that many syncs following its own add.
export function armNextShapeDelay(deck: FakePresentation, syncs: number): void {
  stateOf(deck).armed = syncs;
}

// The hook ShapeCollectionProxy.addTable calls right after creating its
// shape: spends the arm, if any, on exactly this one id.
export function markIfArmed(deck: FakePresentation, shapeId: string): void {
  const state = stateOf(deck);
  if (state.armed === null) return;
  state.pending.set(shapeId, state.armed);
  state.armed = null;
}

// The hook FakeContext.sync() calls once a batch has actually landed: ages
// every blocked id by one, then activates whatever is pending - fresh, never
// aged by the very sync that confirmed its own add.
export function onSyncSucceeded(deck: FakePresentation): void {
  const state = decks.get(deck);
  if (!state) return;
  for (const [id, left] of state.blockedFor) {
    if (left <= 1) state.blockedFor.delete(id);
    else state.blockedFor.set(id, left - 1);
  }
  for (const [id, syncs] of state.pending) state.blockedFor.set(id, syncs);
  state.pending.clear();
}

export function isBlocked(deck: FakePresentation, shapeId: string): boolean {
  return (decks.get(deck)?.blockedFor.get(shapeId) ?? 0) > 0;
}

// shapes.getItem(id), and any proxy re-anchored to it, on PowerPoint for the
// web in this window: code 5010's own message - the fake has no numeric
// codes, so InvalidArgument carries it instead.
export function requireRegistered(
  deck: FakePresentation,
  shapeId: string,
): void {
  if (!isBlocked(deck, shapeId)) return;
  throw Object.assign(new Error("InvalidParam passed to GetItem(id)"), {
    code: "InvalidArgument",
  });
}
