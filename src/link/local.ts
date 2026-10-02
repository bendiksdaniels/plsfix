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

// The floor pushPayload counts nextLocalRev from, given whether this link's
// floor has already been trusted once THIS session (link-record.ts's own
// per-session memory, reset by nothing but a fresh module load - a reload
// after an upgrade, a reopened workbook's pane starting cold, or a Save As
// copy's own separate pane that never ran this session at all). A floor of
// 0 (nothing has ever been pushed for this link) always passes through: no
// deck can hold a rev for an id only now being created. Otherwise, the
// FIRST time this session touches the floor, it is not trusted blind -
// entry.rev/entry.localRev are exactly as good as what this session has
// itself observed so far, which on a first touch is nothing - so it is
// raised to the clock instead of restarting at LOCAL_REV_BASE + 1 or resuming
// whatever local-looking number the registry happened to carry in, either
// of which can reproduce or undercut a revision some deck already holds
// under different content: a registry saved before entry.localRev existed,
// one that never saved a local push before an unsaved reopen, or a Save As
// copy pushed after the original moved on. Any later real moment outweighs
// a counter that only ever climbed by ones from near zero, so a session with
// no memory of its own still lands above every local rev this link has ever
// used, on any device, without needing to remember any of them. Once
// trusted, every later push this session resumes climbing by exactly one,
// unaffected - entry.localRev (link-record.ts) already carries a relay
// excursion within one session correctly, and this never overrides that.
export function localFloor(floor: number, trustedThisSession: boolean): number {
  if (floor === 0 || trustedThisSession) return floor;
  return Math.max(floor, LOCAL_REV_BASE + Date.now());
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
