// In-memory stand-in for the Rust relay: the same rules as server/src/store.rs
// and server/src/relay.rs, minus TTLs and HTTP. Every TypeScript integration
// test runs against this, so a rule that drifts from the server here is a
// production bug the suite cannot see.

import type {
  FetchQuery,
  FetchResult,
  InboxRow,
  OmittedReason,
  RelayApi,
  StatusQuery,
} from "../src/link/relay";
import { FETCH_BLOB_CAP, RelayError } from "../src/link/relay";
import type { RelayStatus } from "../src/link/status";

interface StoredRev {
  rev: number;
  blob: Uint8Array;
}

interface StoredLink extends StoredRev {
  auth: string;
  pushedAt: number;
  // The store keeps two revisions, so the fake keeps the one a push displaced
  // and nothing older: that is exactly what a revert can still reach.
  previous?: StoredRev;
}
interface StoredInbox {
  ws: string;
  id: string;
  auth: string;
  createdAt: number;
  blob: Uint8Array;
}

// server/src/relay.rs refuses a longer batch on both batch routes with
// "400 too many items", so a client that stopped chunking fails here too
// instead of quietly passing against a fake with no ceiling.
const MAX_BATCH_ITEMS = 200;

function refuseOversizedBatch(route: string, items: unknown[]): void {
  if (items.length > MAX_BATCH_ITEMS)
    throw new RelayError(
      "server",
      `relay POST /api/links/${route}: 400 too many items`,
      400,
    );
}

function inboxKey(ws: string, id: string, auth: string): string {
  return `${ws}/${id}/${auth}`;
}

// Why a link a batch asked for carries no blob at all, in the store's order:
// unknown, then a key that does not own it, then a deck that already holds the
// head. Null means a picture is coming.
function omittedReason(
  link: StoredLink | undefined,
  auth: string,
  knownRev: number | undefined,
): OmittedReason | null {
  if (!link) return "missing";
  if (link.auth !== auth) return "auth";
  if (link.rev === knownRev) return "unchanged";
  return null;
}

// Same rules as server/src/store.rs, minus TTLs: first PUT fixes the auth,
// later calls with another auth are forbidden, unknown ids are missing.
export class FakeRelay implements RelayApi {
  readonly links = new Map<string, StoredLink>();
  readonly inbox = new Map<string, StoredInbox>();
  now = 1_000_000;

  async putLink(
    id: string,
    auth: string,
    blob: Uint8Array,
  ): Promise<{ rev: number }> {
    const current = this.links.get(id);
    if (current && current.auth !== auth)
      throw new RelayError("auth", "forbidden", 403);
    const rev = (current?.rev ?? 0) + 1;
    const link: StoredLink = { auth, rev, pushedAt: this.now, blob };
    if (current) link.previous = { rev: current.rev, blob: current.blob };
    this.links.set(id, link);
    return { rev };
  }
  async getLink(
    id: string,
    auth: string,
    knownRev?: number,
  ): Promise<{ rev: number; blob: Uint8Array } | "unchanged"> {
    const current = this.links.get(id);
    if (!current) throw new RelayError("missing", "not found", 404);
    if (current.auth !== auth) throw new RelayError("auth", "forbidden", 403);
    if (knownRev !== undefined && knownRev === current.rev) return "unchanged";
    return { rev: current.rev, blob: current.blob };
  }
  async getLinkRev(
    id: string,
    auth: string,
    rev: number,
  ): Promise<{ rev: number; blob: Uint8Array }> {
    const current = this.links.get(id);
    if (!current) throw new RelayError("missing", "not found", 404);
    if (current.auth !== auth) throw new RelayError("auth", "forbidden", 403);
    const held = [current, current.previous].find(
      (candidate) => candidate?.rev === rev,
    );
    // A revision the two-revision rule has dropped reads like an unknown link.
    if (!held) throw new RelayError("missing", "not found", 404);
    return { rev: held.rev, blob: held.blob };
  }
  async deleteLink(id: string, auth: string): Promise<void> {
    const current = this.links.get(id);
    if (!current) throw new RelayError("missing", "not found", 404);
    if (current.auth !== auth) throw new RelayError("auth", "forbidden", 403);
    this.links.delete(id);
  }
  async status(items: StatusQuery[]): Promise<RelayStatus[]> {
    refuseOversizedBatch("status", items);
    return items.map(({ id, auth }) => {
      const current = this.links.get(id);
      if (!current) return { id, rev: null, pushedAt: null };
      if (current.auth !== auth)
        return { id, rev: null, pushedAt: null, error: "auth" };
      return { id, rev: current.rev, pushedAt: current.pushedAt };
    });
  }
  // Same rules as server/src/fetch.rs: the changed blobs in request order
  // while the cap allows, and every other link named with the reason it
  // carries none. Once one blob does not fit, the rest of the batch is
  // deferred with it rather than sieved.
  async fetchLinks(items: FetchQuery[]): Promise<FetchResult> {
    refuseOversizedBatch("fetch", items);
    const result: FetchResult = { items: [], omitted: [] };
    let room: number | null = FETCH_BLOB_CAP;
    for (const { id, auth, knownRev } of items) {
      const current = this.links.get(id);
      const reason = omittedReason(current, auth, knownRev);
      if (reason !== null) {
        result.omitted.push({ id, reason });
      } else if (current && room !== null && current.blob.length <= room) {
        room -= current.blob.length;
        result.items.push({ id, rev: current.rev, blob: current.blob });
      } else {
        room = null;
        result.omitted.push({ id, reason: "deferred" });
      }
    }
    return result;
  }
  // Keyed like inbox_v2 in server/src/store.rs: (ws, id, auth_hash), never
  // (ws, id). A foreign key that posts the same link id writes its own row
  // beside the pane's item instead of over it, and never sees it - the whole
  // reason that migration exists.
  async postInbox(
    ws: string,
    auth: string,
    id: string,
    blob: Uint8Array,
  ): Promise<void> {
    this.inbox.set(inboxKey(ws, id, auth), {
      ws,
      id,
      auth,
      createdAt: this.now,
      blob,
    });
  }
  async listInbox(ws: string, auth: string): Promise<InboxRow[]> {
    return [...this.inbox.values()]
      .filter((row) => row.ws === ws && row.auth === auth)
      .map((row) => ({
        id: row.id,
        createdAt: row.createdAt,
        blob: row.blob,
      }));
  }
  // Only the key that wrote a row can remove it; another key's delete finds
  // nothing, exactly as the server's WHERE auth_hash = ? does.
  async deleteInbox(ws: string, auth: string, id: string): Promise<void> {
    this.inbox.delete(inboxKey(ws, id, auth));
  }
}
