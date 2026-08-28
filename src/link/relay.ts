// Relay client: a typed fetch wrapper for the /api/links and /api/inbox
// routes, mapping every non-success HTTP outcome (and a fetch rejection) to
// one RelayError kind so callers branch on `.kind` instead of status codes.
// Owns wire shape only; auth, encryption and retry policy live above this.

import { fromBase64Url } from "./crypto";
import type { RelayStatus } from "./status";

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
  deleteLink(id: string, auth: string): Promise<void>;
  status(items: StatusQuery[]): Promise<RelayStatus[]>;
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

  async putLink(
    id: string,
    auth: string,
    blob: Uint8Array,
  ): Promise<{ rev: number }> {
    const response = await this.request(
      `links/${id}`,
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
    const body = (await response.json()) as { rev: number };
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
    const etag = response.headers.get("ETag") ?? "";
    const rev = Number(etag.replace(/^"|"$/g, ""));
    if (!Number.isFinite(rev)) {
      throw new RelayError(
        "server",
        `relay GET ${this.resolve(path).pathname}: bad ETag`,
      );
    }
    const blob = new Uint8Array(await response.arrayBuffer());
    return { rev, blob };
  }

  async deleteLink(id: string, auth: string): Promise<void> {
    await this.request(
      `links/${id}`,
      { method: "DELETE", headers: { Authorization: bearer(auth) } },
      [200, 204],
    );
  }

  async status(items: StatusQuery[]): Promise<RelayStatus[]> {
    const response = await this.request(
      "links/status",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(items),
      },
      [200],
    );
    return (await response.json()) as RelayStatus[];
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
    const response = await this.request(
      `inbox/${ws}`,
      { method: "GET", headers: { Authorization: bearer(auth) } },
      [200],
    );
    const rows = (await response.json()) as {
      id: string;
      createdAt: number;
      blob: string;
    }[];
    return rows.map((row) => ({
      id: row.id,
      createdAt: row.createdAt,
      blob: fromBase64Url(row.blob),
    }));
  }

  async deleteInbox(ws: string, auth: string, id: string): Promise<void> {
    await this.request(
      `inbox/${ws}/${id}`,
      { method: "DELETE", headers: { Authorization: bearer(auth) } },
      [200, 204],
    );
  }
}
