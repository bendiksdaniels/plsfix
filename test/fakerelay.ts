import type { InboxRow, RelayApi, StatusQuery } from "../src/link/relay";
import { RelayError } from "../src/link/relay";
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
  auth: string;
  createdAt: number;
  blob: Uint8Array;
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
    return items.map(({ id, auth }) => {
      const current = this.links.get(id);
      if (!current) return { id, rev: null, pushedAt: null };
      if (current.auth !== auth)
        return { id, rev: null, pushedAt: null, error: "auth" };
      return { id, rev: current.rev, pushedAt: current.pushedAt };
    });
  }
  async postInbox(
    ws: string,
    auth: string,
    id: string,
    blob: Uint8Array,
  ): Promise<void> {
    this.inbox.set(`${ws}/${id}`, { ws, auth, createdAt: this.now, blob });
  }
  async listInbox(ws: string, auth: string): Promise<InboxRow[]> {
    return [...this.inbox.entries()]
      .filter(([, row]) => row.ws === ws && row.auth === auth)
      .map(([key, row]) => ({
        id: key.slice(ws.length + 1),
        createdAt: row.createdAt,
        blob: row.blob,
      }));
  }
  async deleteInbox(ws: string, _auth: string, id: string): Promise<void> {
    this.inbox.delete(`${ws}/${id}`);
  }
}
