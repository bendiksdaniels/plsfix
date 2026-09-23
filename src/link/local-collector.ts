// src/link/local-collector.ts
// Excel's side of local mode: a RelayApi that records what one export or copy
// would have sent, so the unchanged flows in src/excel fill a bundle instead
// of the network. Invariant: it never reads; a read here is a programming error.

import type { Bundle, BundleInboxRow, BundleLink } from "./bundle";
import { nextLocalRev } from "./local";
import type {
  FetchQuery,
  FetchResult,
  InboxRow,
  RelayApi,
  StatusQuery,
  TouchQuery,
} from "./relay";
import type { RelayStatus } from "./status";

function notHere(method: string): Error {
  return new Error(`LocalCollector: ${method} is not used in Excel`);
}

export class LocalCollector implements RelayApi {
  private readonly links = new Map<string, BundleLink>();
  private readonly inbox = new Map<string, BundleInboxRow>();

  constructor(private readonly now: () => number = Date.now) {}

  async putLink(
    id: string,
    _auth: string,
    blob: Uint8Array,
    currentRev = 0,
  ): Promise<{ rev: number }> {
    const held = this.links.get(id)?.rev ?? 0;
    const rev = nextLocalRev(Math.max(currentRev, held));
    this.links.set(id, { id, rev, sentAt: this.now(), blob });
    return { rev };
  }

  async postInbox(_ws: string, _auth: string, id: string, blob: Uint8Array): Promise<void> {
    this.inbox.set(id, { id, createdAt: this.now(), blob });
  }

  // No relay copy to revoke: the deck keeps its object, and the next copy
  // simply leaves the removed link out.
  async deleteLink(_id: string, _auth: string): Promise<void> {}

  async touchLinks(items: TouchQuery[]): Promise<number> {
    return items.length;
  }

  bundle(): Bundle {
    return { links: [...this.links.values()], inbox: [...this.inbox.values()] };
  }

  isEmpty(): boolean {
    return this.links.size === 0 && this.inbox.size === 0;
  }

  async getLink(_id: string, _auth: string, _knownRev?: number): Promise<{ rev: number; blob: Uint8Array } | "unchanged"> {
    throw notHere("getLink");
  }
  async getLinkRev(_id: string, _auth: string, _rev: number): Promise<{ rev: number; blob: Uint8Array }> {
    throw notHere("getLinkRev");
  }
  async status(_items: StatusQuery[]): Promise<RelayStatus[]> {
    throw notHere("status");
  }
  async fetchLinks(_items: FetchQuery[]): Promise<FetchResult> {
    throw notHere("fetchLinks");
  }
  async listInbox(_ws: string, _auth: string): Promise<InboxRow[]> {
    throw notHere("listInbox");
  }
  async deleteInbox(_ws: string, _auth: string, _id: string): Promise<void> {
    throw notHere("deleteInbox");
  }
}
