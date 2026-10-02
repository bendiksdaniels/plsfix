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

  async postInbox(
    _ws: string,
    _auth: string,
    id: string,
    blob: Uint8Array,
  ): Promise<void> {
    this.inbox.set(id, { id, createdAt: this.now(), blob });
  }

  // No relay copy to revoke: the deck keeps its object, and the next copy
  // simply leaves the removed link out. Both arguments are the interface's,
  // never this store's: nothing here reads them.
  async deleteLink(id: string, auth: string): Promise<void> {
    void id;
    void auth;
  }

  async touchLinks(items: TouchQuery[]): Promise<number> {
    return items.length;
  }

  bundle(): Bundle {
    return { links: [...this.links.values()], inbox: [...this.inbox.values()] };
  }

  isEmpty(): boolean {
    return this.links.size === 0 && this.inbox.size === 0;
  }

  // Everything below is a read: a programming error on this side, kept only
  // to satisfy RelayApi. Every parameter is the interface's, unused by design.
  async getLink(
    id: string,
    auth: string,
    knownRev?: number,
  ): Promise<{ rev: number; blob: Uint8Array } | "unchanged"> {
    void id;
    void auth;
    void knownRev;
    throw notHere("getLink");
  }
  async getLinkRev(
    id: string,
    auth: string,
    rev: number,
  ): Promise<{ rev: number; blob: Uint8Array }> {
    void id;
    void auth;
    void rev;
    throw notHere("getLinkRev");
  }
  async status(items: StatusQuery[]): Promise<RelayStatus[]> {
    void items;
    throw notHere("status");
  }
  async fetchLinks(items: FetchQuery[]): Promise<FetchResult> {
    void items;
    throw notHere("fetchLinks");
  }
  async listInbox(ws: string, auth: string): Promise<InboxRow[]> {
    void ws;
    void auth;
    throw notHere("listInbox");
  }
  async deleteInbox(ws: string, auth: string, id: string): Promise<void> {
    void ws;
    void auth;
    void id;
    throw notHere("deleteInbox");
  }
}
