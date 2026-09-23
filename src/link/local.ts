// src/link/local.ts
// Local mode's shared rules: the revision space above the relay's, Revert's
// previous rev in either space, and the fixed workspace that seals inbox rows
// in a bundle. Invariant: a local rev never equals a relay rev.

import { deriveWorkspace, type Workspace } from "./workspace";

// Relay revs are small integers the server counts; local revs live above 2^40,
// so a switch in either direction always reads as an update (the deck compares
// revs by inequality) and never as "up to date" by accident.
export const LOCAL_REV_BASE = 2 ** 40;

export function isLocalRev(rev: number): boolean {
  return rev > LOCAL_REV_BASE;
}

// The rev after whatever the registry holds, from either space.
export function nextLocalRev(current: number): number {
  return Math.max(current, LOCAL_REV_BASE) + 1;
}

// One below, never under the first rev of its own space.
export function previousRevOf(rev: number): number | null {
  const floor = isLocalRev(rev) ? LOCAL_REV_BASE + 1 : 1;
  const previous = rev - 1;
  return Number.isInteger(previous) && previous >= floor ? previous : null;
}

// NOT a secret, on purpose: every device derives it from the same all-zero
// secret. It seals a bundle's inbox rows only so both transports share one
// envelope and listInbox / insertFromInbox run unchanged; the bundle itself
// is what must stay private.
let local: Promise<Workspace> | null = null;

export function localWorkspace(): Promise<Workspace> {
  local ??= deriveWorkspace(new Uint8Array(32));
  return local;
}
