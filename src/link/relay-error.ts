// The one failure type every relay call fails with, and the status-to-kind
// map behind it, so a caller branches on `.kind` instead of on HTTP codes.
// Split out of relay.ts to keep both files under the length limit; the calls
// live there, the body shapes their answers are checked against in wire.ts.
// Pure - no fetch, no Office.js.
// Invariant: a kind added to the union is added to RELAY_ERROR_KINDS in the
// same edit, or the build stops.

export type RelayErrorKind =
  "network" | "auth" | "missing" | "tooLarge" | "server" | "timeout";

export class RelayError extends Error {
  readonly kind: RelayErrorKind;
  readonly status: number | undefined;

  constructor(kind: RelayErrorKind, message: string, status?: number) {
    super(message);
    this.name = "RelayError";
    this.kind = kind;
    this.status = status;
  }
}

// instanceof fails across module graphs - a test that resets modules, or two
// bundles each holding their own copy - so a RelayError is recognised by its
// name plus a kind the union knows. The record keeps the two in step: a kind
// added to the union and not listed here stops the build.
const RELAY_ERROR_KINDS: Record<RelayErrorKind, true> = {
  network: true,
  auth: true,
  missing: true,
  tooLarge: true,
  server: true,
  timeout: true,
};

export function isRelayError(error: unknown): error is RelayError {
  if (!(error instanceof Error) || error.name !== "RelayError") return false;
  const kind: unknown = (error as { readonly kind?: unknown }).kind;
  return typeof kind === "string" && Object.hasOwn(RELAY_ERROR_KINDS, kind);
}

// The relay's refusals in kinds a caller can act on. The two limit refusals -
// 429 (asked too often) and 507 (store full) - are neither the caller's key
// nor its payload, so they land in "server" like any other fault the pane can
// only report; the relay's own reason travels in the message beside them.
export function statusKind(status: number): RelayErrorKind {
  if (status === 401 || status === 403) return "auth";
  if (status === 404) return "missing";
  if (status === 413) return "tooLarge";
  return "server";
}
