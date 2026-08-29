// The shapes the relay routes promise, and the guards that check a 200's body
// against them. The relay is the untrusted half of this design (its path
// carries an Access bypass), so a JSON body is checked here instead of cast:
// relay.ts turns anything that fails into a typed RelayError, and no caller
// ever sees a raw SyntaxError. Pure - no fetch, no Office.js.

import type { RelayStatus } from "./status";

// `"3"` from the relay, `W/"3"` once an intermediary (Cloudflare weakens a
// strong tag whenever it rewrites a body) has been through it. Anything else -
// including a missing header, which Number("") would have made rev 0 - is a
// bad response, never a revision.
export const ETAG_REV = /^(?:W\/)?"(\d+)"$/;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isNumberOrNull(value: unknown): boolean {
  return value === null || typeof value === "number";
}

export function isPutResult(value: unknown): value is { rev: number } {
  return isRecord(value) && typeof value.rev === "number";
}

// Why the relay refused, in the shape every non-success body carries:
// {"error":"too many items"}. Without it a 400 reaches the pane as a bare
// status code, where "your batch is too long", "your id is malformed" and a
// genuine server fault all read the same.
export function isErrorBody(value: unknown): value is { error: string } {
  return isRecord(value) && typeof value.error === "string";
}

export function isStatusRow(value: unknown): value is RelayStatus {
  return (
    isRecord(value) &&
    typeof value.id === "string" &&
    isNumberOrNull(value.rev) &&
    isNumberOrNull(value.pushedAt) &&
    (value.error === undefined || value.error === "auth")
  );
}

export interface InboxJson {
  id: string;
  createdAt: number;
  blob: string;
}

export function isInboxJson(value: unknown): value is InboxJson {
  return (
    isRecord(value) &&
    typeof value.id === "string" &&
    typeof value.createdAt === "number" &&
    typeof value.blob === "string"
  );
}

// Why a link the batch fetch was asked for carries no blob. The first three
// are final answers; "deferred" only means "not in this response", and the
// caller asks for that link on its own.
export type OmittedReason = "unchanged" | "missing" | "auth" | "deferred";

// The record keeps the union honest: a reason added to one and not the other
// stops the build.
const OMITTED_REASONS: Record<OmittedReason, true> = {
  unchanged: true,
  missing: true,
  auth: true,
  deferred: true,
};

export interface FetchJson {
  id: string;
  rev: number;
  blob: string;
}

export interface OmittedJson {
  id: string;
  reason: OmittedReason;
}

export interface FetchJsonBody {
  items: FetchJson[];
  omitted: OmittedJson[];
}

function isFetchJson(value: unknown): value is FetchJson {
  return (
    isRecord(value) &&
    typeof value.id === "string" &&
    typeof value.rev === "number" &&
    typeof value.blob === "string"
  );
}

function isOmittedJson(value: unknown): value is OmittedJson {
  return (
    isRecord(value) &&
    typeof value.id === "string" &&
    typeof value.reason === "string" &&
    Object.hasOwn(OMITTED_REASONS, value.reason)
  );
}

// A batch answer is two lists, and a body missing either of them is not one:
// a client that read `items` off a half-shaped body would report the rest of
// the deck as up to date.
export function isFetchJsonBody(value: unknown): value is FetchJsonBody {
  return (
    isRecord(value) &&
    arrayOf(isFetchJson)(value.items) &&
    arrayOf(isOmittedJson)(value.omitted)
  );
}

export function arrayOf<T>(
  guard: (value: unknown) => value is T,
): (value: unknown) => value is T[] {
  return (value: unknown): value is T[] =>
    Array.isArray(value) && value.every(guard);
}
