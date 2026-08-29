// Relay client: a typed fetch wrapper for the /api/links and /api/inbox
// routes, mapping every non-success HTTP outcome (and a fetch rejection) to
// one RelayError kind so callers branch on `.kind` instead of status codes.
// Owns the calls; the body shapes and their guards live in wire.ts, and auth,
// encryption and retry policy live above this.
// Invariant: every failure leaves this module as a RelayError - a 200 whose
// body is not the JSON shape the route promises is one too, never a raw parse
// error, so the pane's toast always has a `kind` to render.

import { fromBase64Url } from "./crypto";
import type { RelayStatus } from "./status";
import {
  arrayOf,
  ETAG_REV,
  isFetchJsonBody,
  isInboxJson,
  isPutResult,
  isStatusRow,
  type OmittedReason,
} from "./wire";

export type { OmittedReason } from "./wire";

export type RelayErrorKind =
  "network" | "auth" | "missing" | "tooLarge" | "server";

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
};

export function isRelayError(error: unknown): error is RelayError {
  if (!(error instanceof Error) || error.name !== "RelayError") return false;
  const kind: unknown = (error as { readonly kind?: unknown }).kind;
  return typeof kind === "string" && Object.hasOwn(RELAY_ERROR_KINDS, kind);
}

export interface StatusQuery {
  id: string;
  auth: string;
}

// The deck's side of one link in a batch fetch: which link, the key that opens
// it, and the revision the deck already holds. Without `knownRev` the relay
// always answers with a blob.
export interface FetchQuery {
  id: string;
  auth: string;
  knownRev?: number;
}

export interface FetchedLink {
  id: string;
  rev: number;
  blob: Uint8Array;
}

export interface OmittedLink {
  id: string;
  reason: OmittedReason;
}

// Every link the batch was asked for lands in exactly one of the two lists.
export interface FetchResult {
  items: FetchedLink[];
  omitted: OmittedLink[];
}

// The relay's per-batch limits, so a caller can chunk before it calls: more
// items than this is a 400, and blobs past the cap come back "deferred".
export const MAX_FETCH_ITEMS = 200;
export const FETCH_BLOB_CAP = 4 * 1024 * 1024;

export interface InboxRow {
  id: string;
  createdAt: number;
  blob: Uint8Array;
}

export interface RelayApi {
  putLink(id: string, auth: string, blob: Uint8Array): Promise<{ rev: number }>;
  getLink(
    id: string,
    auth: string,
    knownRev?: number,
  ): Promise<{ rev: number; blob: Uint8Array } | "unchanged">;
  getLinkRev(
    id: string,
    auth: string,
    rev: number,
  ): Promise<{ rev: number; blob: Uint8Array }>;
  deleteLink(id: string, auth: string): Promise<void>;
  status(items: StatusQuery[]): Promise<RelayStatus[]>;
  fetchLinks(items: FetchQuery[]): Promise<FetchResult>;
  postInbox(
    ws: string,
    auth: string,
    id: string,
    blob: Uint8Array,
  ): Promise<void>;
  listInbox(ws: string, auth: string): Promise<InboxRow[]>;
  deleteInbox(ws: string, auth: string, id: string): Promise<void>;
}

// The pane's document URL always ends in a page (taskpane.html, pptpane.html);
// the relay sits beside it as a sibling "api/" directory, one level up.
export function relayBaseUrl(documentUrl: string): URL {
  return new URL("api/", documentUrl);
}

function bearer(auth: string): string {
  return `Bearer ${auth}`;
}

function statusKind(status: number): RelayErrorKind {
  if (status === 401 || status === 403) return "auth";
  if (status === 404) return "missing";
  if (status === 413) return "tooLarge";
  return "server";
}

// BodyInit requires the buffer generic pinned to ArrayBuffer; .slice() copies
// into a fresh ArrayBuffer-backed view, which a plain Uint8Array parameter
// (buffer typed ArrayBufferLike) does not guarantee.
function toBody(bytes: Uint8Array): Uint8Array<ArrayBuffer> {
  return bytes.slice();
}

export class RelayClient implements RelayApi {
  private readonly baseUrl: URL;
  private readonly fetchImpl: typeof fetch;

  constructor(baseUrl: URL | string, fetchImpl?: typeof fetch) {
    this.baseUrl = typeof baseUrl === "string" ? new URL(baseUrl) : baseUrl;
    this.fetchImpl = fetchImpl ?? globalThis.fetch.bind(globalThis);
  }

  private resolve(path: string): URL {
    return new URL(path, this.baseUrl);
  }

  // Every call funnels through here so status-to-kind mapping happens once:
  // a response whose status is in `expect` is handed back untouched, anything
  // else becomes a RelayError, and a rejected fetch becomes kind "network".
  private async request(
    path: string,
    init: RequestInit,
    expect: number[],
  ): Promise<Response> {
    const url = this.resolve(path);
    const method = init.method ?? "GET";
    let response: Response;
    try {
      response = await this.fetchImpl(url.href, init);
    } catch {
      throw new RelayError(
        "network",
        `relay ${method} ${url.pathname}: network error`,
      );
    }
    if (expect.includes(response.status)) return response;
    throw new RelayError(
      statusKind(response.status),
      `relay ${method} ${url.pathname}: ${response.status}`,
      response.status,
    );
  }

  // One message for every way a 200 can still be unusable: not JSON at all (a
  // captive portal, an Access interstitial, the dev proxy answering with
  // index.html), the wrong shape, or a blob that is not base64url.
  private badResponse(method: string, path: string): RelayError {
    return new RelayError(
      "server",
      `relay ${method} ${this.resolve(path).pathname}: bad response`,
    );
  }

  private async json<T>(
    response: Response,
    method: string,
    path: string,
    guard: (value: unknown) => value is T,
  ): Promise<T> {
    let parsed: unknown;
    try {
      parsed = await response.json();
    } catch {
      throw this.badResponse(method, path);
    }
    if (!guard(parsed)) throw this.badResponse(method, path);
    return parsed;
  }

  async putLink(
    id: string,
    auth: string,
    blob: Uint8Array,
  ): Promise<{ rev: number }> {
    const path = `links/${id}`;
    const response = await this.request(
      path,
      {
        method: "PUT",
        headers: {
          Authorization: bearer(auth),
          "Content-Type": "application/octet-stream",
        },
        body: toBody(blob),
      },
      [200],
    );
    const body = await this.json(response, "PUT", path, isPutResult);
    return { rev: body.rev };
  }

  async getLink(
    id: string,
    auth: string,
    knownRev?: number,
  ): Promise<{ rev: number; blob: Uint8Array } | "unchanged"> {
    const path = `links/${id}`;
    const headers: Record<string, string> = { Authorization: bearer(auth) };
    if (knownRev !== undefined) headers["If-None-Match"] = `"${knownRev}"`;
    const response = await this.request(
      path,
      { method: "GET", headers },
      [200, 304],
    );
    if (response.status === 304) return "unchanged";
    return this.blobFrom(response, path);
  }

  // One exact revision, the relay's second-newest at most: what "Revert last
  // update" repaints from. No If-None-Match, because the deck is asking for a
  // picture it does not hold; a revision retention or the TTL has dropped is
  // a 404 like any other missing link.
  async getLinkRev(
    id: string,
    auth: string,
    rev: number,
  ): Promise<{ rev: number; blob: Uint8Array }> {
    const path = `links/${id}?rev=${String(rev)}`;
    const response = await this.request(
      path,
      { method: "GET", headers: { Authorization: bearer(auth) } },
      [200],
    );
    return this.blobFrom(response, path);
  }

  // The revision a body belongs to is the ETag's word, never the caller's
  // guess, so a relay that answered with another one is a bad response.
  private async blobFrom(
    response: Response,
    path: string,
  ): Promise<{ rev: number; blob: Uint8Array }> {
    const match = ETAG_REV.exec(response.headers.get("ETag") ?? "");
    if (match === null) {
      throw new RelayError(
        "server",
        `relay GET ${this.resolve(path).pathname}: bad ETag`,
      );
    }
    return {
      rev: Number(match[1]),
      blob: new Uint8Array(await response.arrayBuffer()),
    };
  }

  async deleteLink(id: string, auth: string): Promise<void> {
    await this.request(
      `links/${id}`,
      { method: "DELETE", headers: { Authorization: bearer(auth) } },
      [200, 204],
    );
  }

  async status(items: StatusQuery[]): Promise<RelayStatus[]> {
    const path = "links/status";
    const response = await this.request(
      path,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(items),
      },
      [200],
    );
    return this.json(response, "POST", path, arrayOf(isStatusRow));
  }

  // Every changed picture of a deck in one round trip. The relay leaves out
  // whatever it will not send (unchanged, missing, foreign, or past its
  // response cap) and names it in `omitted`, so a caller can tell "nothing to
  // do" from "ask again on your own" without counting rows.
  async fetchLinks(items: FetchQuery[]): Promise<FetchResult> {
    const path = "links/fetch";
    const response = await this.request(
      path,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(items),
      },
      [200],
    );
    const body = await this.json(response, "POST", path, isFetchJsonBody);
    return {
      items: body.items.map((item) => ({
        id: item.id,
        rev: item.rev,
        blob: this.decode(item.blob, "POST", path),
      })),
      omitted: body.omitted,
    };
  }

  async postInbox(
    ws: string,
    auth: string,
    id: string,
    blob: Uint8Array,
  ): Promise<void> {
    await this.request(
      `inbox/${ws}`,
      {
        method: "POST",
        headers: {
          Authorization: bearer(auth),
          "Content-Type": "application/octet-stream",
          "X-SMT-Link-Id": id,
        },
        body: toBody(blob),
      },
      [200],
    );
  }

  async listInbox(ws: string, auth: string): Promise<InboxRow[]> {
    const path = `inbox/${ws}`;
    const response = await this.request(
      path,
      { method: "GET", headers: { Authorization: bearer(auth) } },
      [200],
    );
    const rows = await this.json(response, "GET", path, arrayOf(isInboxJson));
    return rows.map((row) => ({
      id: row.id,
      createdAt: row.createdAt,
      blob: this.decode(row.blob, "GET", path),
    }));
  }

  // fromBase64Url throws a plain Error on anything outside the alphabet; one
  // bad blob must not leave a listing or a batch untyped.
  private decode(blob: string, method: string, path: string): Uint8Array {
    try {
      return fromBase64Url(blob);
    } catch {
      throw this.badResponse(method, path);
    }
  }

  async deleteInbox(ws: string, auth: string, id: string): Promise<void> {
    await this.request(
      `inbox/${ws}/${id}`,
      { method: "DELETE", headers: { Authorization: bearer(auth) } },
      [200, 204],
    );
  }
}
