// The relay's refusals in the pane's words: one sentence per answer a modeller
// can act on, so a 429 or a 507 never reaches a toast as "relay GET
// /api/links: 429". Pure - no fetch, no Office.js; the kinds themselves are
// relay-error.ts, and a kind added there with a sentence worth saying is added
// here in the same edit.
// Invariant: an answer with no sentence of its own returns undefined, and the
// relay's own message travels on unchanged.

import { isRelayError } from "./relay-error";

// A 413 is the relay's 4 MiB link route, or an nginx or Cloudflare hop in
// front of it; "413 payload too large" tells a modeller nothing they can act
// on. Excel refuses the same export with its own copy of this sentence.
export const TOO_LARGE =
  "That export is too big to send. Export a smaller range from Excel.";
export const RELAY_BUSY = "The relay is busy. Try again in a minute.";
export const RELAY_FULL = "The relay is full. Ask for space to be cleared.";
export const RELAY_FAULT = "The relay had a problem. Try again in a minute.";

// server/src/relay_gates.rs: 429 past the per-client rate limit, 507 past the
// byte ceiling or the inbox row cap.
const TOO_MANY = 429;
const NO_SPACE = 507;
const SERVER_FAULT = 500;

export function relayReason(error: unknown): string | undefined {
  if (!isRelayError(error)) return undefined;
  if (error.kind === "tooLarge") return TOO_LARGE;
  const status = error.status;
  if (status === undefined) return undefined;
  if (status === TOO_MANY) return RELAY_BUSY;
  if (status === NO_SPACE) return RELAY_FULL;
  return status >= SERVER_FAULT ? RELAY_FAULT : undefined;
}
