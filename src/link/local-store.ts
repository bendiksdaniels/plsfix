// src/link/local-store.ts
// PowerPoint's side of local mode: a RelayApi answered from the bundles pasted
// on this computer, so Update, Revert, Insert, Change source and Paste latest
// linked run unchanged. Keeps the two highest revisions per link, as the relay
// keeps two. Invariant: an older bundle never downgrades a link; no network.

import type { Bundle, BundleLink } from "./bundle";
import type {
  LocalPersistence,
  StoredLocalInbox,
  StoredLocalLink,
} from "./local-persist";
import {
  RelayError,
  type FetchQuery,
  type FetchResult,
  type InboxRow,
  type RelayApi,
  type StatusQuery,
  type TouchQuery,
} from "./relay";
import type { RelayStatus } from "./status";

export const LOCAL_STORE_BUDGET = 32 * 1024 * 1024;
export const NOT_PASTED_HERE =
  "No copy of this link was pasted on this computer. Copy it again in Excel.";

export interface IngestResult {
  linkIds: string[];
  waiting: number;
}

// The two highest revisions win, whatever order the bundles arrive in.
export function mergeRevision(
  held: StoredLocalLink | undefined,
  incoming: BundleLink,
  pastedAt: number,
): StoredLocalLink {
  const { id, rev, sentAt, blob } = incoming;
  if (held === undefined) return { id, rev, sentAt, blob, pastedAt };
  if (rev === held.rev) return { ...held, pastedAt };
  if (rev > held.rev) {
    const previous = { rev: held.rev, sentAt: held.sentAt, blob: held.blob };
    return { id, rev, sentAt, blob, pastedAt, previous };
  }
  if (held.previous === undefined || rev > held.previous.rev) {
    return { ...held, pastedAt, previous: { rev, sentAt, blob } };
  }
  return { ...held, pastedAt };
}

function linkBytes(link: StoredLocalLink): number {
  return link.blob.length + (link.previous?.blob.length ?? 0);
}

function notHere(method: string): Error {
  return new Error(
    `LocalStore: ${method} is Excel's; local mode writes through ingest()`,
  );
}

export class LocalStore implements RelayApi {
  private readonly links = new Map<string, StoredLocalLink>();
  private readonly inbox = new Map<string, StoredLocalInbox>();

  private constructor(
    private readonly persist: LocalPersistence,
    private readonly now: () => number,
    private readonly budget: number,
  ) {}

  static async open(
    persist: LocalPersistence,
    options: { now?: () => number; budget?: number } = {},
  ): Promise<LocalStore> {
    const store = new LocalStore(
      persist,
      options.now ?? Date.now,
      options.budget ?? LOCAL_STORE_BUDGET,
    );
    const saved = await persist.load();
    for (const link of saved.links) store.links.set(link.id, link);
    for (const row of saved.inbox) store.inbox.set(row.id, row);
    return store;
  }

  durable(): boolean {
    return this.persist.durable;
  }

  async ingest(
    bundle: Bundle,
    deckIds: ReadonlySet<string>,
  ): Promise<IngestResult> {
    const pastedAt = this.now();
    const links = bundle.links.map((link) => {
      const next = mergeRevision(this.links.get(link.id), link, pastedAt);
      this.links.set(link.id, next);
      return next;
    });
    const rows = bundle.inbox
      .filter((row) => !deckIds.has(row.id))
      .map((row): StoredLocalInbox => ({ ...row, pastedAt }));
    for (const row of rows) this.inbox.set(row.id, row);
    const evicted = this.evict(new Set(bundle.links.map((link) => link.id)));
    await this.persist.putLinks(
      links.filter((link) => this.links.has(link.id)),
    );
    await this.persist.putInbox(rows.filter((row) => this.inbox.has(row.id)));
    await this.persist.deleteLinks(evicted.links);
    await this.persist.deleteInbox(evicted.inbox);
    return {
      linkIds: bundle.links.map((link) => link.id),
      waiting: rows.length,
    };
  }

  async clear(): Promise<void> {
    this.links.clear();
    this.inbox.clear();
    await this.persist.clear();
  }

  // Oldest paste first, never a link of the paste in progress: whole links,
  // then inbox rows, until the blobs fit the budget.
  private evict(keep: ReadonlySet<string>): {
    links: string[];
    inbox: string[];
  } {
    let total = 0;
    for (const link of this.links.values()) total += linkBytes(link);
    for (const row of this.inbox.values()) total += row.blob.length;
    const out = { links: [] as string[], inbox: [] as string[] };
    const oldLinks = [...this.links.values()]
      .filter((link) => !keep.has(link.id))
      .sort((a, b) => a.pastedAt - b.pastedAt);
    for (const link of oldLinks) {
      if (total <= this.budget) break;
      total -= linkBytes(link);
      this.links.delete(link.id);
      out.links.push(link.id);
    }
    const oldRows = [...this.inbox.values()]
      .filter((row) => !keep.has(row.id))
      .sort((a, b) => a.pastedAt - b.pastedAt);
    for (const row of oldRows) {
      if (total <= this.budget) break;
      total -= row.blob.length;
      this.inbox.delete(row.id);
      out.inbox.push(row.id);
    }
    return out;
  }

  async status(items: StatusQuery[]): Promise<RelayStatus[]> {
    return items.map(({ id }) => {
      const link = this.links.get(id);
      return link
        ? { id, rev: link.rev, pushedAt: link.sentAt }
        : { id, rev: null, pushedAt: null, local: true };
    });
  }

  async fetchLinks(items: FetchQuery[]): Promise<FetchResult> {
    const result: FetchResult = { items: [], omitted: [] };
    for (const { id, knownRev } of items) {
      const link = this.links.get(id);
      if (link === undefined) result.omitted.push({ id, reason: "notPasted" });
      else if (link.rev === knownRev)
        result.omitted.push({ id, reason: "unchanged" });
      else result.items.push({ id, rev: link.rev, blob: link.blob });
    }
    return result;
  }

  async getLink(
    id: string,
    _auth: string,
    knownRev?: number,
  ): Promise<{ rev: number; blob: Uint8Array } | "unchanged"> {
    const link = this.require(id);
    if (link.rev === knownRev) return "unchanged";
    return { rev: link.rev, blob: link.blob };
  }

  async getLinkRev(
    id: string,
    _auth: string,
    rev: number,
  ): Promise<{ rev: number; blob: Uint8Array }> {
    const link = this.require(id);
    if (link.rev === rev) return { rev, blob: link.blob };
    if (link.previous?.rev === rev) return { rev, blob: link.previous.blob };
    throw new RelayError("missing", NOT_PASTED_HERE);
  }

  // Local mode has one workspace (src/link/local.ts), so both arguments are
  // the interface's only: every pasted row answers, whoever asks.
  async listInbox(ws: string, auth: string): Promise<InboxRow[]> {
    void ws;
    void auth;
    return [...this.inbox.values()]
      .sort((a, b) => b.createdAt - a.createdAt)
      .map(({ id, createdAt, blob }) => ({ id, createdAt, blob }));
  }

  async deleteInbox(ws: string, auth: string, id: string): Promise<void> {
    void ws;
    void auth;
    if (this.inbox.delete(id)) await this.persist.deleteInbox([id]);
  }

  // Everything below is a write: Excel's, never this deck's. Every parameter
  // is the interface's, unused by design.
  async putLink(
    id: string,
    auth: string,
    blob: Uint8Array,
    currentRev?: number,
  ): Promise<{ rev: number }> {
    void id;
    void auth;
    void blob;
    void currentRev;
    throw notHere("putLink");
  }
  async postInbox(
    ws: string,
    auth: string,
    id: string,
    blob: Uint8Array,
  ): Promise<void> {
    void ws;
    void auth;
    void id;
    void blob;
    throw notHere("postInbox");
  }
  async deleteLink(id: string, auth: string): Promise<void> {
    void id;
    void auth;
    throw notHere("deleteLink");
  }
  async touchLinks(items: TouchQuery[]): Promise<number> {
    void items;
    throw notHere("touchLinks");
  }

  private require(id: string): StoredLocalLink {
    const link = this.links.get(id);
    if (link === undefined) throw new RelayError("missing", NOT_PASTED_HERE);
    return link;
  }
}
