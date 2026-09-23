# Local links Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Linked objects move from Excel to PowerPoint on one computer by one copy and one paste, with no relay; the relay stays as a per-device switch.

**Architecture:** Two new implementations of the existing `RelayApi` seam (`src/link/relay.ts`): a `LocalCollector` in Excel that records what an export or copy would have sent, and a `LocalStore` in PowerPoint that answers every deck flow from pasted bundles (persisted in IndexedDB). A bundle travels in the clipboard's HTML flavor; the plain-text flavor is one sentence. Every existing flow in `src/excel/` and `src/ppt/` runs unchanged against either transport.

**Tech Stack:** TypeScript (strict), Vite, Vitest (node + jsdom), Office.js, WebCrypto, IndexedDB (`fake-indexeddb` in tests), Playwright-core for the browser probe.

**Spec:** `docs/superpowers/specs/2026-09-23-local-links-design.md` (read it first; this plan argues from it).

## Global Constraints

- Every source file opens with a 2-4 line header comment: purpose, what it owns, its key invariant.
- File cap 400 lines, function cap 50 lines. New logic goes in NEW files. `src/ppt/main.ts` (620),
  `src/pane/links-tab.ts` (538) and `src/ppt/links.ts` (436) are already over the cap: they may only
  gain wiring lines, never logic.
- `test/fakehost.ts` and `test/fakeppt/*` are additive-only: never change an existing line.
- No `it.skip`, `it.todo`, `describe.skip`. Vitest count only goes up.
- Every existing assertion you change is a user-visible wording change: list each one old -> new in
  your report.
- Help sentences (`src/help/copy.*.ts`): at most 140 characters, no em dash, end with a full stop;
  every new `<button>` needs one (`src/help/copy.test.ts` enforces it both ways).
- Pane copy is English; the manual (`manual/src/content/*.rs`) is Latvian (normal Latvian, no
  "cilne", "lietotne", "darbplūsma").
- Exact values from the spec: `LOCAL_REV_BASE = 2 ** 40`; bundle cap 25 MB (`25 * 1024 * 1024`
  characters of JSON); store budget 32 MB (`32 * 1024 * 1024` blob bytes); storage key
  `plsfix.link.transport.v1`; IndexedDB database `plsfix-links`, stores `links` and `inbox`;
  clipboard attribute `data-plsfix-links`; plain-text sentence
  "pls,fix links for PowerPoint: paste them in the pls,fix pane, Inbox tab."
- Gates before a slice is reported done: `npm run check`, `npm run ux:sweep -- --port <own>`,
  `npm run ux:check -- --port <own>`, all green. Ports per worktree are given in each brief.
- Commits: terse imperative subject, 3-6 line body naming the files and the behaviour change, no
  Claude trailer. Do not merge, tag or push.

## File structure

| File | Slice | Owns |
|---|---|---|
| `src/link/bundle.ts` (new) | S1 | the bundle JSON codec, the HTML carrier, reading a paste |
| `src/link/local.ts` (new) | S1 | `LOCAL_REV_BASE`, the next local rev, Revert's previous rev, the fixed local workspace |
| `src/link/local-collector.ts` (new) | S1 | Excel's `RelayApi` that records one action into a bundle |
| `src/link/local-persist.ts` (new) | S1 | the persistence port, memory and IndexedDB adapters |
| `src/link/local-store.ts` (new) | S1 | PowerPoint's `RelayApi` answered from pasted bundles |
| `src/link/transport-setting.ts` (new) | S1 | the per-device `local` / `relay` setting and its default rule |
| `src/ui/clipboard-links.ts` (new) | S1 | the two-flavor clipboard write and the paste read |
| `src/link/status.ts`, `src/link/wire.ts`, `src/ppt/fetch.ts`, `src/ppt/links.ts`, `src/ppt/views.ts`, `src/styles/ppt.css`, `pptpane.html` | S1 | the `notPasted` status end to end |
| `src/link/relay.ts`, `src/excel/link-record.ts` | S1 | `putLink`'s optional `currentRev` |
| `src/ppt/revert.ts` | S1 | Revert's previous rev via `previousRevOf`, the neutral "no longer available" sentence |
| `test/relay-contract.ts` (new) | S1 | the read rules FakeRelay and LocalStore share |
| `src/excel/links.ts` | S2 | `pushLinks(ids, relay, { announce })` |
| `src/pane/links-transport.ts` (new) | S2 | the Excel side of the setting, `copyLocal()`, the Copy for PowerPoint bar |
| `src/pane/links-tab.ts`, `taskpane.html`, `src/styles/links.css` or the sheet that owns the Links tab | S2 | wiring, the new section, labels, visibility |
| `src/ppt/transport.ts` (new) | S3 | which `RelayApi` and workspace the deck flows use; the Settings select |
| `src/ppt/paste-links.ts` (new) | S3 | `pasteLinks()` and its summary line |
| `src/ppt/main.ts`, `pptpane.html`, `src/styles/ppt.css` | S3 | wiring, the paste box, Clear pasted links, the "Sent" column |
| help copy, README, INSTALL, privacy page, manual, Map, `scripts/ux/clipboard-probe.mjs` | S4 | the words and the permanent proofs |

---

## Slice S1: the core (no visible change in relay mode except two neutral sentences)

### Task 1: The bundle codec and its clipboard carrier

**Files:**
- Create: `src/link/bundle.ts`
- Test: `src/link/bundle.test.ts`

**Interfaces:**
- Consumes: `toBase64Url`, `fromBase64Url` (`src/link/crypto.ts`), `isLinkId` (`src/link/model.ts`).
- Produces: `interface BundleLink { id: string; rev: number; sentAt: number; blob: Uint8Array }`,
  `interface BundleInboxRow { id: string; createdAt: number; blob: Uint8Array }`,
  `interface Bundle { links: BundleLink[]; inbox: BundleInboxRow[] }`,
  `type BundleRead = { ok: true; bundle: Bundle } | { ok: false; reason: "notBundle" | "newerVersion" }`,
  `encodeBundle(bundle: Bundle): string`, `decodeBundle(text: string): BundleRead`,
  `bundleHtml(json: string): string`, `readPastedBundle(html: string, plain: string): BundleRead`,
  constants `BUNDLE_VERSION = 1`, `BUNDLE_MAX_CHARS`, `BUNDLE_SENTENCE`.

- [ ] **Step 1: Write the failing tests**

```ts
// src/link/bundle.test.ts
// The clipboard bundle: a round trip, the guards foreign text meets, and the
// HTML carrier a paste reads first.
import { describe, expect, it } from "vitest";
import {
  BUNDLE_SENTENCE,
  bundleHtml,
  decodeBundle,
  encodeBundle,
  readPastedBundle,
  type Bundle,
} from "./bundle";

const ID = "0123456789abcdef0123456789abcdef";
const sample: Bundle = {
  links: [{ id: ID, rev: 2 ** 40 + 1, sentAt: 1_790_000_000_000, blob: new Uint8Array([1, 2, 3]) }],
  inbox: [{ id: ID, createdAt: 1_790_000_000_000, blob: new Uint8Array([9, 8]) }],
};

describe("encodeBundle / decodeBundle", () => {
  it("round-trips every field", () => {
    const read = decodeBundle(encodeBundle(sample));
    expect(read).toEqual({ ok: true, bundle: sample });
  });

  it("answers notBundle for text that is not one", () => {
    for (const text of ["", "hello", "{}", "[1]", '{"plsfix":"other","v":1}']) {
      expect(decodeBundle(text)).toEqual({ ok: false, reason: "notBundle" });
    }
  });

  it("answers newerVersion for a bundle from a newer add-in", () => {
    const newer = encodeBundle(sample).replace('"v":1', '"v":2');
    expect(decodeBundle(newer)).toEqual({ ok: false, reason: "newerVersion" });
  });

  it("refuses a bad id, a bad rev or a blob that is not base64url", () => {
    const text = encodeBundle(sample);
    expect(decodeBundle(text.replace(ID, "not-an-id")).ok).toBe(false);
    expect(decodeBundle(text.replace(`"rev":${String(2 ** 40 + 1)}`, '"rev":-1')).ok).toBe(false);
    expect(decodeBundle(text.replace('"blob":"AQID"', '"blob":"@@"')).ok).toBe(false);
  });

  it("refuses an empty bundle", () => {
    expect(decodeBundle(encodeBundle({ links: [], inbox: [] })).ok).toBe(false);
  });
});

describe("the clipboard carrier", () => {
  it("hides the bundle in one attribute and shows only the sentence", () => {
    const html = bundleHtml(encodeBundle(sample));
    expect(html.startsWith('<span data-plsfix-links="')).toBe(true);
    expect(html.endsWith(`">${BUNDLE_SENTENCE}</span>`)).toBe(true);
  });

  it("reads the HTML flavor first, the plain text second", () => {
    const json = encodeBundle(sample);
    expect(readPastedBundle(bundleHtml(json), BUNDLE_SENTENCE)).toEqual({ ok: true, bundle: sample });
    expect(readPastedBundle("", `  ${json}\n`)).toEqual({ ok: true, bundle: sample });
    expect(readPastedBundle("<b>hi</b>", "hi")).toEqual({ ok: false, reason: "notBundle" });
  });

  it("survives a sanitizer that re-serialises the span around it", () => {
    const json = encodeBundle(sample);
    const wrapped = `<html><body><!--StartFragment-->${bundleHtml(json).replace("<span ", '<span style="color:black" ')}<!--EndFragment--></body></html>`;
    expect(readPastedBundle(wrapped, "").ok).toBe(true);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

Run: `npx vitest run src/link/bundle.test.ts`
Expected: FAIL, "Failed to resolve import ./bundle".

- [ ] **Step 3: Implement**

```ts
// src/link/bundle.ts
// The clipboard bundle: everything PowerPoint needs to insert or update a set
// of links, carried by one copy in Excel and one paste in PowerPoint. Pure:
// the JSON codec and the two clipboard flavors, no DOM, no Office.js.
// Invariant: foreign text never throws here; it answers a reason instead.

import { fromBase64Url, toBase64Url } from "./crypto";
import { isLinkId } from "./model";

export const BUNDLE_VERSION = 1;
// The JSON is ASCII (hex ids, base64url blobs, numbers), so chars = bytes.
export const BUNDLE_MAX_CHARS = 25 * 1024 * 1024;
export const BUNDLE_SENTENCE =
  "pls,fix links for PowerPoint: paste them in the pls,fix pane, Inbox tab.";
// base64url never holds a quote, so the attribute value needs no escaping and
// a sanitizer that re-serialises the span leaves it byte for byte.
const CARRIER = /data-plsfix-links="([A-Za-z0-9_-]+)"/;

export interface BundleLink {
  id: string;
  rev: number;
  sentAt: number;
  blob: Uint8Array;
}

export interface BundleInboxRow {
  id: string;
  createdAt: number;
  blob: Uint8Array;
}

export interface Bundle {
  links: BundleLink[];
  inbox: BundleInboxRow[];
}

export type BundleRead =
  | { ok: true; bundle: Bundle }
  | { ok: false; reason: "notBundle" | "newerVersion" };

const NOT_BUNDLE: BundleRead = { ok: false, reason: "notBundle" };

export function encodeBundle(bundle: Bundle): string {
  return JSON.stringify({
    plsfix: "links",
    v: BUNDLE_VERSION,
    links: bundle.links.map((link) => ({
      id: link.id,
      rev: link.rev,
      sentAt: link.sentAt,
      blob: toBase64Url(link.blob),
    })),
    inbox: bundle.inbox.map((row) => ({
      id: row.id,
      createdAt: row.createdAt,
      blob: toBase64Url(row.blob),
    })),
  });
}

export function decodeBundle(text: string): BundleRead {
  let value: unknown;
  try {
    value = JSON.parse(text);
  } catch {
    return NOT_BUNDLE;
  }
  if (!isRecord(value) || value.plsfix !== "links") return NOT_BUNDLE;
  if (typeof value.v !== "number") return NOT_BUNDLE;
  if (value.v > BUNDLE_VERSION) return { ok: false, reason: "newerVersion" };
  if (value.v !== BUNDLE_VERSION) return NOT_BUNDLE;
  try {
    const links = listOf(value.links, readLink);
    const inbox = listOf(value.inbox, readRow);
    if (links === null || inbox === null) return NOT_BUNDLE;
    if (links.length === 0 && inbox.length === 0) return NOT_BUNDLE;
    return { ok: true, bundle: { links, inbox } };
  } catch {
    return NOT_BUNDLE;
  }
}

export function bundleHtml(json: string): string {
  const carried = toBase64Url(new TextEncoder().encode(json));
  return `<span data-plsfix-links="${carried}">${BUNDLE_SENTENCE}</span>`;
}

// What a paste carries: the HTML flavor the pane wrote first, then plain text
// (a bundle copied by hand out of Excel's fallback box).
export function readPastedBundle(html: string, plain: string): BundleRead {
  const carried = CARRIER.exec(html)?.[1];
  if (carried !== undefined) {
    try {
      return decodeBundle(new TextDecoder().decode(fromBase64Url(carried)));
    } catch {
      return NOT_BUNDLE;
    }
  }
  return decodeBundle(plain.trim());
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

// fromBase64Url throws on a bad blob; decodeBundle turns that into notBundle.
function listOf<T>(value: unknown, read: (item: unknown) => T | null): T[] | null {
  if (!Array.isArray(value)) return null;
  const out: T[] = [];
  for (const item of value) {
    const read1 = read(item);
    if (read1 === null) return null;
    out.push(read1);
  }
  return out;
}

function readLink(item: unknown): BundleLink | null {
  if (!isRecord(item)) return null;
  const { id, rev, sentAt, blob } = item;
  if (typeof id !== "string" || !isLinkId(id)) return null;
  if (typeof rev !== "number" || !Number.isSafeInteger(rev) || rev < 1) return null;
  if (typeof sentAt !== "number" || !Number.isFinite(sentAt)) return null;
  if (typeof blob !== "string" || blob === "") return null;
  return { id, rev, sentAt, blob: fromBase64Url(blob) };
}

function readRow(item: unknown): BundleInboxRow | null {
  if (!isRecord(item)) return null;
  const { id, createdAt, blob } = item;
  if (typeof id !== "string" || !isLinkId(id)) return null;
  if (typeof createdAt !== "number" || !Number.isFinite(createdAt)) return null;
  if (typeof blob !== "string" || blob === "") return null;
  return { id, createdAt, blob: fromBase64Url(blob) };
}
```

- [ ] **Step 4: Run to verify it passes**

Run: `npx vitest run src/link/bundle.test.ts`
Expected: PASS (8 tests).

- [ ] **Step 5: Commit**

```bash
git add src/link/bundle.ts src/link/bundle.test.ts
git commit -m "Link bundle: the JSON codec and its clipboard carrier" -m "src/link/bundle.ts: encode/decode a bundle of sealed links and inbox rows, the HTML span that carries it, the paste reader (HTML first, plain text second). Pure; foreign text answers a reason, never throws."
```

### Task 2: Local revisions and the fixed local workspace

**Files:**
- Create: `src/link/local.ts`
- Test: `src/link/local.test.ts`

**Interfaces:**
- Consumes: `deriveWorkspace`, `type Workspace` (`src/link/workspace.ts`).
- Produces: `LOCAL_REV_BASE`, `isLocalRev(rev: number): boolean`, `nextLocalRev(current: number): number`,
  `previousRevOf(rev: number): number | null`, `localWorkspace(): Promise<Workspace>`.

- [ ] **Step 1: Write the failing tests**

```ts
// src/link/local.test.ts
// The local revision space and the fixed workspace every device derives.
import { describe, expect, it } from "vitest";
import { isLocalRev, LOCAL_REV_BASE, localWorkspace, nextLocalRev, previousRevOf } from "./local";

describe("local revisions", () => {
  it("start above the relay's space and count up from there", () => {
    expect(nextLocalRev(0)).toBe(LOCAL_REV_BASE + 1);
    expect(nextLocalRev(7)).toBe(LOCAL_REV_BASE + 1);
    expect(nextLocalRev(LOCAL_REV_BASE + 4)).toBe(LOCAL_REV_BASE + 5);
  });

  it("tell the two spaces apart", () => {
    expect(isLocalRev(12)).toBe(false);
    expect(isLocalRev(LOCAL_REV_BASE + 1)).toBe(true);
  });

  it("give Revert the rev below, never under the first of its space", () => {
    expect(previousRevOf(5)).toBe(4);
    expect(previousRevOf(1)).toBeNull();
    expect(previousRevOf(LOCAL_REV_BASE + 3)).toBe(LOCAL_REV_BASE + 2);
    expect(previousRevOf(LOCAL_REV_BASE + 1)).toBeNull();
  });
});

describe("localWorkspace", () => {
  it("is the same on every call and every device", async () => {
    const a = await localWorkspace();
    const b = await localWorkspace();
    expect(a.id).toBe(b.id);
    expect(a.id).toMatch(/^[A-Za-z0-9_-]+$/);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

Run: `npx vitest run src/link/local.test.ts`
Expected: FAIL, "Failed to resolve import ./local".

- [ ] **Step 3: Implement**

```ts
// src/link/local.ts
// Local mode's shared rules: the revision space above the relay's, Revert's
// previous rev in either space, and the fixed workspace that seals inbox rows
// in a bundle. Invariant: a local rev never equals a relay rev.

import { deriveWorkspace, type Workspace } from "./workspace";

// Relay revs are small integers the server counts; local revs live above 2^40,
// so a switch in either direction always reads as an update (the deck compares
// revs by inequality) and never as "up to date" by accident.
export const LOCAL_REV_BASE = 2 ** 40;

export function isLocalRev(rev: number): boolean {
  return rev > LOCAL_REV_BASE;
}

// The rev after whatever the registry holds, from either space.
export function nextLocalRev(current: number): number {
  return Math.max(current, LOCAL_REV_BASE) + 1;
}

// One below, never under the first rev of its own space.
export function previousRevOf(rev: number): number | null {
  const floor = isLocalRev(rev) ? LOCAL_REV_BASE + 1 : 1;
  const previous = rev - 1;
  return Number.isInteger(previous) && previous >= floor ? previous : null;
}

// NOT a secret, on purpose: every device derives it from the same all-zero
// secret. It seals a bundle's inbox rows only so both transports share one
// envelope and listInbox / insertFromInbox run unchanged; the bundle itself
// is what must stay private.
let local: Promise<Workspace> | null = null;

export function localWorkspace(): Promise<Workspace> {
  local ??= deriveWorkspace(new Uint8Array(32));
  return local;
}
```

- [ ] **Step 4: Run to verify it passes**

Run: `npx vitest run src/link/local.test.ts`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add src/link/local.ts src/link/local.test.ts
git commit -m "Local revs above 2^40 and the fixed local workspace" -m "src/link/local.ts: nextLocalRev, previousRevOf and isLocalRev keep local and relay revs apart; localWorkspace is derived from a fixed secret so bundle inbox rows share the relay's envelope."
```

### Task 3: `putLink` learns the registry's rev; the LocalCollector

**Files:**
- Modify: `src/link/relay.ts` (the `RelayApi.putLink` line and `RelayClient.putLink`'s signature only)
- Modify: `src/excel/link-record.ts:44` (`pushPayload` passes `entry.rev`)
- Create: `src/link/local-collector.ts`
- Test: `src/link/local-collector.test.ts`

**Interfaces:**
- Consumes: `nextLocalRev` (Task 2), `Bundle`, `BundleLink`, `BundleInboxRow` (Task 1).
- Produces: `RelayApi.putLink(id: string, auth: string, blob: Uint8Array, currentRev?: number): Promise<{ rev: number }>`;
  `class LocalCollector implements RelayApi` with `constructor(now?: () => number)`, `bundle(): Bundle`,
  `isEmpty(): boolean`.

- [ ] **Step 1: Write the failing tests**

```ts
// src/link/local-collector.test.ts
// What one Excel action records instead of sending.
import { describe, expect, it } from "vitest";
import { LOCAL_REV_BASE } from "./local";
import { LocalCollector } from "./local-collector";

const ID = "0123456789abcdef0123456789abcdef";

describe("LocalCollector", () => {
  it("answers the next local rev after the registry's", async () => {
    const collector = new LocalCollector(() => 1000);
    expect(await collector.putLink(ID, "auth", new Uint8Array([1]), 3)).toEqual({ rev: LOCAL_REV_BASE + 1 });
    expect(await collector.putLink(ID, "auth", new Uint8Array([2]), LOCAL_REV_BASE + 1)).toEqual({ rev: LOCAL_REV_BASE + 2 });
  });

  it("keeps the last put per link and every inbox row", async () => {
    const collector = new LocalCollector(() => 1000);
    await collector.putLink(ID, "a", new Uint8Array([1]));
    await collector.putLink(ID, "a", new Uint8Array([2]));
    await collector.postInbox("ws", "auth", ID, new Uint8Array([7]));
    expect(collector.bundle()).toEqual({
      links: [{ id: ID, rev: LOCAL_REV_BASE + 2, sentAt: 1000, blob: new Uint8Array([2]) }],
      inbox: [{ id: ID, createdAt: 1000, blob: new Uint8Array([7]) }],
    });
  });

  it("has nothing to revoke or touch, and refuses every read", async () => {
    const collector = new LocalCollector();
    await expect(collector.deleteLink(ID, "a")).resolves.toBeUndefined();
    await expect(collector.touchLinks([{ id: ID, auth: "a" }])).resolves.toBe(1);
    await expect(collector.status([])).rejects.toThrow(/not used in Excel/);
    await expect(collector.listInbox("ws", "a")).rejects.toThrow(/not used in Excel/);
    expect(collector.isEmpty()).toBe(true);
  });
});
```

- [ ] **Step 2: Run to verify it fails**

Run: `npx vitest run src/link/local-collector.test.ts`
Expected: FAIL, "Failed to resolve import ./local-collector".

- [ ] **Step 3: Change the interface, then implement**

In `src/link/relay.ts`, the `RelayApi` interface line becomes:

```ts
  putLink(
    id: string,
    auth: string,
    blob: Uint8Array,
    // The registry's rev before this push. The relay counts its own and ignores
    // it; the local collector counts on from it (src/link/local.ts).
    currentRev?: number,
  ): Promise<{ rev: number }>;
```

`RelayClient.putLink` keeps its body and gains the unused optional parameter only if the compiler
asks for it (an implementation may take fewer parameters). In `src/excel/link-record.ts` line 44:

```ts
  return (await relay.putLink(entry.id, keys.auth, blob, entry.rev)).rev;
```

Then:

```ts
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
```

- [ ] **Step 4: Run to verify it passes, and nothing else moved**

Run: `npx vitest run src/link/local-collector.test.ts && npx tsc --noEmit`
Expected: PASS (3 tests); tsc clean.

- [ ] **Step 5: Commit**

```bash
git add src/link/relay.ts src/excel/link-record.ts src/link/local-collector.ts src/link/local-collector.test.ts
git commit -m "LocalCollector: record an export or copy instead of sending it" -m "src/link/local-collector.ts: a RelayApi that keeps the last put per link and every inbox row as a bundle. putLink takes the registry's rev (optional; the relay ignores it) so local revs count on from it."
```

### Task 4: The `notPasted` status, end to end

**Files:**
- Modify: `src/link/status.ts` (`LinkStatus`, `RelayStatus`, `deriveStatus`)
- Modify: `src/link/wire.ts` (`OmittedReason` + `OMITTED_REASONS`)
- Modify: `src/ppt/fetch.ts` (`FetchOutcome`, `emptyOutcome`, `countOmitted`)
- Modify: `src/ppt/links.ts` (`UpdateSummary`, the fold at ~line 183, `summarize`)
- Modify: `src/ppt/views.ts` (`STATUS_LABELS`, `STATUS_CLASSES`)
- Modify: `src/styles/ppt.css` (`.badge.notpasted` next to `.badge.missing` at ~line 253, colour from an existing neutral token only)
- Modify: `pptpane.html` (status filter: `<option value="notPasted">Not pasted yet</option>` after "Wrong link key")
- Test: `src/link/status.test.ts`, `src/ppt/fetch.test.ts`, `src/ppt/links.test.ts` (or the file that already tests `summarize`)

**Interfaces:**
- Produces: `LinkStatus` gains `"notPasted"`; `RelayStatus` gains `local?: true`; `OmittedReason` gains
  `"notPasted"`; `FetchOutcome.notPasted: number`; `UpdateSummary.notPasted: number`.

- [ ] **Step 1: Write the failing tests**

```ts
// added to src/link/status.test.ts
describe("deriveStatus in local mode", () => {
  it("says notPasted when this computer holds no copy", () => {
    expect(deriveStatus(5, { id: "x", rev: null, pushedAt: null, local: true })).toBe("notPasted");
  });
  it("says current when the deck already holds something newer than this computer", () => {
    expect(deriveStatus(2 ** 40 + 5, { id: "x", rev: 2 ** 40 + 3, pushedAt: 1, local: true })).toBe("current");
  });
  it("keeps the relay's rules: missing without a rev, any inequality is an update", () => {
    expect(deriveStatus(5, { id: "x", rev: null, pushedAt: null })).toBe("missing");
    expect(deriveStatus(5, { id: "x", rev: 1, pushedAt: 1 })).toBe("updateAvailable");
  });
});
```

```ts
// added to the test file that covers summarize (grep "summarize(" under src/ppt)
it("counts links this computer holds no copy of", () => {
  expect(summarize({ updated: 1, current: 0, notPasted: 2, missing: 0, wrongKey: 0, failed: 0, sourceChanges: [], failures: [], notes: [] }))
    .toBe("1 updated, 2 not pasted yet");
});
```

For `fetch.ts`, add one case to its existing test file: a stub relay whose `fetchLinks` answers
`{ items: [], omitted: [{ id, reason: "notPasted" }] }` for one row; expect `outcome.notPasted === 1`
and `outcome.missing === 0`. Build the row with the helpers that file already uses.

- [ ] **Step 2: Run to verify they fail**

Run: `npx vitest run src/link/status.test.ts src/ppt`
Expected: FAIL (type errors or wrong status).

- [ ] **Step 3: Implement**

`src/link/status.ts`:

```ts
export type LinkStatus =
  | "current"
  | "updateAvailable"
  | "missing"
  | "wrongKey"
  | "notPasted";

export interface RelayStatus {
  id: string;
  rev: number | null;
  pushedAt: number | null;
  error?: "auth";
  // Answered by this computer's LocalStore, not a relay (local mode).
  local?: true;
}

export function deriveStatus(
  tagRev: number,
  relay: RelayStatus | undefined,
): LinkStatus {
  if (relay === undefined) return "missing";
  if (relay.error === "auth") return "wrongKey";
  if (relay.rev === null) return relay.local ? "notPasted" : "missing";
  // This computer's copy is older than what the deck holds: nothing newer here.
  if (relay.local && relay.rev < tagRev) return "current";
  // Any inequality, not just a higher rev: a link swept at its 7-day TTL comes
  // back from the next push as rev 1, so a relay rev *below* the tag's is the
  // ordinary "the deck is stale" case, not an impossibility.
  if (relay.rev !== tagRev) return "updateAvailable";
  return "current";
}
```

`src/link/wire.ts`: add `| "notPasted"` to `OmittedReason` with a comment "(local mode only: this
computer holds no copy)", and `notPasted: true` to `OMITTED_REASONS`.

`src/ppt/fetch.ts`: `FetchOutcome` gains `notPasted: number`; `emptyOutcome()` sets it to 0;
`countOmitted` gains `else if (reason === "notPasted") outcome.notPasted += group.rows.length;` before
the final `else`.

`src/ppt/links.ts`: `UpdateSummary` gains `notPasted: number`; every place that builds an
`UpdateSummary` literal sets it to 0 (the compiler lists them); the fold near line 183 gains
`summary.notPasted += fetched.notPasted;`; `summarize` gains `[summary.notPasted, "not pasted yet"]`
right after `"up to date"`.

`src/ppt/views.ts`: `notPasted: "Not pasted yet"` in `STATUS_LABELS`, `notPasted: "badge notpasted"`
in `STATUS_CLASSES`.

- [ ] **Step 4: Run to verify they pass**

Run: `npx vitest run src/link src/ppt test/ppt*.ts && npx tsc --noEmit`
Expected: PASS; tsc clean.

- [ ] **Step 5: Commit**

```bash
git add -A src/link/status.ts src/link/status.test.ts src/link/wire.ts src/ppt pptpane.html src/styles/ppt.css
git commit -m "notPasted: the status of a link this computer holds no copy of" -m "src/link/status.ts, wire.ts, src/ppt/fetch.ts, links.ts, views.ts, ppt.css, pptpane.html: a local answer with no rev reads 'Not pasted yet' (badge, filter, summary); a local copy older than the deck reads up to date. Relay rules unchanged."
```

### Task 5: The persistence port and the LocalStore

**Files:**
- Create: `src/link/local-persist.ts` (the port and `memoryPersistence()`; Task 6 adds IndexedDB)
- Create: `src/link/local-store.ts`
- Create: `test/relay-contract.ts`
- Test: `src/link/local-store.test.ts`, `test/relay-contract.integration.test.ts`

**Interfaces:**
- Consumes: `Bundle`, `BundleLink` (Task 1), `RelayError` (`src/link/relay-error.ts`, re-exported from `./relay`), `RelayStatus` (Task 4).
- Produces:
  - `interface StoredLocalLink { id: string; rev: number; sentAt: number; blob: Uint8Array; pastedAt: number; previous?: { rev: number; sentAt: number; blob: Uint8Array } }`
  - `interface StoredLocalInbox { id: string; createdAt: number; blob: Uint8Array; pastedAt: number }`
  - `interface LocalPersistence { readonly durable: boolean; load(): Promise<{ links: StoredLocalLink[]; inbox: StoredLocalInbox[] }>; putLinks(links: StoredLocalLink[]): Promise<void>; putInbox(rows: StoredLocalInbox[]): Promise<void>; deleteLinks(ids: string[]): Promise<void>; deleteInbox(ids: string[]): Promise<void>; clear(): Promise<void> }`
  - `memoryPersistence(): LocalPersistence`
  - `class LocalStore implements RelayApi` with `static open(persist: LocalPersistence, options?: { now?: () => number; budget?: number }): Promise<LocalStore>`, `ingest(bundle: Bundle, deckIds: ReadonlySet<string>): Promise<IngestResult>`, `clear(): Promise<void>`, `durable(): boolean`
  - `interface IngestResult { linkIds: string[]; waiting: number }`, `mergeRevision(held, incoming, pastedAt): StoredLocalLink`
  - `NOT_PASTED_HERE` (the sentence), `LOCAL_STORE_BUDGET`
  - `relayReadContract(name: string, make: () => Promise<ContractTransport>): void` in `test/relay-contract.ts`

- [ ] **Step 1: Write the shared contract and the store's own tests**

```ts
// test/relay-contract.ts
// The read rules every RelayApi the PowerPoint pane runs against must keep:
// FakeRelay (the server's rules) and LocalStore (local mode). A rule kept by
// one and not the other is a transport the deck cannot trust.
import { describe, expect, it } from "vitest";
import { isRelayError, type RelayApi } from "../src/link/relay";
import { deriveStatus } from "../src/link/status";

export interface ContractTransport {
  relay: RelayApi;
  auth: string;
  // Stores the blobs as consecutive revisions of one link, oldest first, and
  // answers the revs it gave them.
  seed(id: string, blobs: Uint8Array[]): Promise<number[]>;
  unknown: "missing" | "notPasted";
}

const ID = "0123456789abcdef0123456789abcdef";
const OTHER = "fedcba9876543210fedcba9876543210";
const [ONE, TWO, THREE] = [1, 2, 3].map((n) => new Uint8Array([n]));

export function relayReadContract(name: string, make: () => Promise<ContractTransport>): void {
  describe(`${name}: the read rules the deck relies on`, () => {
    it("answers the newest revision in status and fetch", async () => {
      const t = await make();
      const [, rev2] = await t.seed(ID, [ONE!, TWO!]);
      const [status] = await t.relay.status([{ id: ID, auth: t.auth }]);
      expect(status?.rev).toBe(rev2);
      const fetched = await t.relay.fetchLinks([{ id: ID, auth: t.auth }]);
      expect(fetched.items).toEqual([{ id: ID, rev: rev2, blob: TWO }]);
    });

    it("omits a link whose newest revision the deck already holds", async () => {
      const t = await make();
      const [rev1] = await t.seed(ID, [ONE!]);
      const fetched = await t.relay.fetchLinks([{ id: ID, auth: t.auth, knownRev: rev1 }]);
      expect(fetched).toEqual({ items: [], omitted: [{ id: ID, reason: "unchanged" }] });
    });

    it("reaches the revision before the newest and nothing older", async () => {
      const t = await make();
      const [rev1, rev2] = await t.seed(ID, [ONE!, TWO!, THREE!]);
      await expect(t.relay.getLinkRev(ID, t.auth, rev2!)).resolves.toEqual({ rev: rev2, blob: TWO });
      const gone = await t.relay.getLinkRev(ID, t.auth, rev1!).catch((error: unknown) => error);
      expect(isRelayError(gone) && gone.kind === "missing").toBe(true);
    });

    it("says why an unknown link has no picture", async () => {
      const t = await make();
      const fetched = await t.relay.fetchLinks([{ id: OTHER, auth: t.auth }]);
      expect(fetched.omitted).toEqual([{ id: OTHER, reason: t.unknown }]);
      const [status] = await t.relay.status([{ id: OTHER, auth: t.auth }]);
      expect(deriveStatus(3, status)).toBe(t.unknown);
    });
  });
}
```

```ts
// test/relay-contract.integration.test.ts
// The same read rules, run on the fake relay and on this computer's store.
import { LOCAL_REV_BASE } from "../src/link/local";
import { memoryPersistence } from "../src/link/local-persist";
import { LocalStore } from "../src/link/local-store";
import { FakeRelay } from "./fakerelay";
import { relayReadContract } from "./relay-contract";

relayReadContract("FakeRelay", async () => {
  const relay = new FakeRelay();
  return {
    relay,
    auth: "auth-1",
    unknown: "missing",
    seed: async (id, blobs) => {
      const revs: number[] = [];
      for (const blob of blobs) revs.push((await relay.putLink(id, "auth-1", blob)).rev);
      return revs;
    },
  };
});

relayReadContract("LocalStore", async () => {
  const store = await LocalStore.open(memoryPersistence());
  return {
    relay: store,
    auth: "any",
    unknown: "notPasted",
    seed: async (id, blobs) => {
      const revs = blobs.map((_, index) => LOCAL_REV_BASE + 1 + index);
      for (const [index, blob] of blobs.entries()) {
        await store.ingest({ links: [{ id, rev: revs[index]!, sentAt: index, blob }], inbox: [] }, new Set());
      }
      return revs;
    },
  };
});
```

```ts
// src/link/local-store.test.ts
// The store's own rules: no downgrade, the Inbox skips what the deck holds,
// the budget evicts the oldest paste, a reopen brings everything back.
import { describe, expect, it } from "vitest";
import type { Bundle } from "./bundle";
import { LOCAL_REV_BASE } from "./local";
import { memoryPersistence, type LocalPersistence, type StoredLocalInbox, type StoredLocalLink } from "./local-persist";
import { LocalStore, NOT_PASTED_HERE } from "./local-store";

const A = "a".repeat(32);
const B = "b".repeat(32);
const R = (n: number): number => LOCAL_REV_BASE + n;
const bytes = (n: number, fill = 1): Uint8Array => new Uint8Array(n).fill(fill);

function bundle(links: [string, number, number][], inbox: string[] = []): Bundle {
  return {
    links: links.map(([id, rev, size]) => ({ id, rev, sentAt: rev, blob: bytes(size, rev % 250) })),
    inbox: inbox.map((id) => ({ id, createdAt: 1, blob: bytes(4) })),
  };
}

// Keeps what it is given, so a second LocalStore.open reads it back.
function mapPersistence(): LocalPersistence {
  const links = new Map<string, StoredLocalLink>();
  const inbox = new Map<string, StoredLocalInbox>();
  return {
    durable: true,
    load: async () => ({ links: [...links.values()], inbox: [...inbox.values()] }),
    putLinks: async (list) => list.forEach((l) => links.set(l.id, l)),
    putInbox: async (list) => list.forEach((r) => inbox.set(r.id, r)),
    deleteLinks: async (ids) => ids.forEach((id) => links.delete(id)),
    deleteInbox: async (ids) => ids.forEach((id) => inbox.delete(id)),
    clear: async () => {
      links.clear();
      inbox.clear();
    },
  };
}

describe("LocalStore.ingest", () => {
  it("never downgrades: an older bundle leaves the newest in place", async () => {
    const store = await LocalStore.open(memoryPersistence());
    await store.ingest(bundle([[A, R(3), 2]]), new Set());
    await store.ingest(bundle([[A, R(1), 2]]), new Set());
    const [status] = await store.status([{ id: A, auth: "x" }]);
    expect(status?.rev).toBe(R(3));
    await expect(store.getLinkRev(A, "x", R(1))).resolves.toEqual({ rev: R(1), blob: bytes(2, R(1) % 250) });
  });

  it("keeps inbox rows only for links the deck does not hold", async () => {
    const store = await LocalStore.open(memoryPersistence());
    const result = await store.ingest(bundle([[A, R(1), 2], [B, R(1), 2]], [A, B]), new Set([A]));
    expect(result).toEqual({ linkIds: [A, B], waiting: 1 });
    expect((await store.listInbox("ws", "x")).map((row) => row.id)).toEqual([B]);
  });

  it("evicts the oldest paste first, never the one in progress", async () => {
    let now = 0;
    const store = await LocalStore.open(memoryPersistence(), { now: () => (now += 1), budget: 10 });
    await store.ingest(bundle([[A, R(1), 6]]), new Set());
    await store.ingest(bundle([[B, R(1), 6]]), new Set());
    const [a, b] = await store.status([{ id: A, auth: "x" }, { id: B, auth: "x" }]);
    expect(a?.rev).toBeNull();
    expect(b?.rev).toBe(R(1));
  });

  it("comes back from its persistence after a reopen, and clear empties both", async () => {
    const persist = mapPersistence();
    const first = await LocalStore.open(persist);
    await first.ingest(bundle([[A, R(2), 2]], [A]), new Set());
    const second = await LocalStore.open(persist);
    expect((await second.status([{ id: A, auth: "x" }]))[0]?.rev).toBe(R(2));
    expect(await second.listInbox("ws", "x")).toHaveLength(1);
    await second.clear();
    expect((await (await LocalStore.open(persist)).status([{ id: A, auth: "x" }]))[0]?.rev).toBeNull();
  });

  it("says in a sentence that a link was never pasted here", async () => {
    const store = await LocalStore.open(memoryPersistence());
    await expect(store.getLink(A, "x")).rejects.toThrow(NOT_PASTED_HERE);
  });
});
```

- [ ] **Step 2: Run to verify they fail**

Run: `npx vitest run src/link/local-store.test.ts test/relay-contract.integration.test.ts`
Expected: FAIL, unresolved imports.

- [ ] **Step 3: Implement the port**

```ts
// src/link/local-persist.ts
// Where pasted links live between sessions: the persistence port LocalStore
// writes through, its memory stand-in, and (Task 6) IndexedDB. Invariant: no
// method throws into a paste; a failed write only turns `durable` false.

export interface StoredLocalLink {
  id: string;
  rev: number;
  sentAt: number;
  blob: Uint8Array;
  pastedAt: number;
  previous?: { rev: number; sentAt: number; blob: Uint8Array };
}

export interface StoredLocalInbox {
  id: string;
  createdAt: number;
  blob: Uint8Array;
  pastedAt: number;
}

export interface LocalPersistence {
  readonly durable: boolean;
  load(): Promise<{ links: StoredLocalLink[]; inbox: StoredLocalInbox[] }>;
  putLinks(links: StoredLocalLink[]): Promise<void>;
  putInbox(rows: StoredLocalInbox[]): Promise<void>;
  deleteLinks(ids: string[]): Promise<void>;
  deleteInbox(ids: string[]): Promise<void>;
  clear(): Promise<void>;
}

// Nothing survives the pane: LocalStore's own maps are the memory.
export function memoryPersistence(): LocalPersistence {
  const nothing = async (): Promise<void> => undefined;
  return {
    durable: false,
    load: async () => ({ links: [], inbox: [] }),
    putLinks: nothing,
    putInbox: nothing,
    deleteLinks: nothing,
    deleteInbox: nothing,
    clear: nothing,
  };
}
```

- [ ] **Step 4: Implement the store**

```ts
// src/link/local-store.ts
// PowerPoint's side of local mode: a RelayApi answered from the bundles pasted
// on this computer, so Update, Revert, Insert, Change source and Paste latest
// linked run unchanged. Keeps the two highest revisions per link, as the relay
// keeps two. Invariant: an older bundle never downgrades a link; no network.

import type { Bundle, BundleLink } from "./bundle";
import type { LocalPersistence, StoredLocalInbox, StoredLocalLink } from "./local-persist";
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
  return new Error(`LocalStore: ${method} is Excel's; local mode writes through ingest()`);
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
    const store = new LocalStore(persist, options.now ?? Date.now, options.budget ?? LOCAL_STORE_BUDGET);
    const saved = await persist.load();
    for (const link of saved.links) store.links.set(link.id, link);
    for (const row of saved.inbox) store.inbox.set(row.id, row);
    return store;
  }

  durable(): boolean {
    return this.persist.durable;
  }

  async ingest(bundle: Bundle, deckIds: ReadonlySet<string>): Promise<IngestResult> {
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
    await this.persist.putLinks(links.filter((link) => this.links.has(link.id)));
    await this.persist.putInbox(rows.filter((row) => this.inbox.has(row.id)));
    await this.persist.deleteLinks(evicted.links);
    await this.persist.deleteInbox(evicted.inbox);
    return { linkIds: bundle.links.map((link) => link.id), waiting: rows.length };
  }

  async clear(): Promise<void> {
    this.links.clear();
    this.inbox.clear();
    await this.persist.clear();
  }

  // Oldest paste first, never a link of the paste in progress: whole links,
  // then inbox rows, until the blobs fit the budget.
  private evict(keep: ReadonlySet<string>): { links: string[]; inbox: string[] } {
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
      else if (link.rev === knownRev) result.omitted.push({ id, reason: "unchanged" });
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

  async getLinkRev(id: string, _auth: string, rev: number): Promise<{ rev: number; blob: Uint8Array }> {
    const link = this.require(id);
    if (link.rev === rev) return { rev, blob: link.blob };
    if (link.previous?.rev === rev) return { rev, blob: link.previous.blob };
    throw new RelayError("missing", NOT_PASTED_HERE);
  }

  async listInbox(_ws: string, _auth: string): Promise<InboxRow[]> {
    return [...this.inbox.values()]
      .sort((a, b) => b.createdAt - a.createdAt)
      .map(({ id, createdAt, blob }) => ({ id, createdAt, blob }));
  }

  async deleteInbox(_ws: string, _auth: string, id: string): Promise<void> {
    if (this.inbox.delete(id)) await this.persist.deleteInbox([id]);
  }

  async putLink(_id: string, _auth: string, _blob: Uint8Array, _currentRev?: number): Promise<{ rev: number }> {
    throw notHere("putLink");
  }
  async postInbox(_ws: string, _auth: string, _id: string, _blob: Uint8Array): Promise<void> {
    throw notHere("postInbox");
  }
  async deleteLink(_id: string, _auth: string): Promise<void> {
    throw notHere("deleteLink");
  }
  async touchLinks(_items: TouchQuery[]): Promise<number> {
    throw notHere("touchLinks");
  }

  private require(id: string): StoredLocalLink {
    const link = this.links.get(id);
    if (link === undefined) throw new RelayError("missing", NOT_PASTED_HERE);
    return link;
  }
}
```

- [ ] **Step 5: Run to verify they pass**

Run: `npx vitest run src/link test/relay-contract.integration.test.ts`
Expected: PASS (5 store tests, 8 contract tests: 4 per transport). FakeRelay answers an unknown id
with `{ rev: null }` in `status` and `missing` in `fetchLinks` (checked 23.09), which the contract expects.

- [ ] **Step 6: Commit**

```bash
git add src/link/local-persist.ts src/link/local-store.ts src/link/local-store.test.ts test/relay-contract.ts test/relay-contract.integration.test.ts
git commit -m "LocalStore: the deck's RelayApi answered from pasted bundles" -m "src/link/local-store.ts keeps the two highest revs per link, inbox rows only for links the deck lacks, a 32 MB budget evicting the oldest paste. test/relay-contract.ts holds FakeRelay and LocalStore to the same read rules."
```

### Task 6: IndexedDB persistence

**Files:**
- Modify: `src/link/local-persist.ts` (add `openLocalPersistence`)
- Modify: `package.json` / `package-lock.json` (`npm install -D fake-indexeddb`)
- Test: `src/link/local-persist.test.ts`

**Interfaces:**
- Produces: `openLocalPersistence(factory?: IDBFactory): Promise<LocalPersistence>` (memory when
  `factory` is missing or refuses to open).

- [ ] **Step 1: Add the dev dependency**

Run: `npm install -D fake-indexeddb`
Expected: `package.json` devDependencies gains `fake-indexeddb`; the lockfile updates.

- [ ] **Step 2: Write the failing tests**

```ts
// src/link/local-persist.test.ts
// IndexedDB behind the port: a round trip, deletes, clear, and memory when
// the webview has no IndexedDB or refuses it.
import { IDBFactory } from "fake-indexeddb";
import { describe, expect, it } from "vitest";
import { openLocalPersistence } from "./local-persist";

const link = { id: "a".repeat(32), rev: 5, sentAt: 1, blob: new Uint8Array([1, 2]), pastedAt: 2 };
const row = { id: "b".repeat(32), createdAt: 3, blob: new Uint8Array([3]), pastedAt: 4 };

describe("openLocalPersistence", () => {
  it("keeps links and inbox rows across a reopen", async () => {
    const factory = new IDBFactory();
    const first = await openLocalPersistence(factory);
    expect(first.durable).toBe(true);
    await first.putLinks([link]);
    await first.putInbox([row]);
    const loaded = await (await openLocalPersistence(factory)).load();
    expect(loaded.links).toEqual([link]);
    expect(loaded.inbox).toEqual([row]);
  });

  it("deletes and clears", async () => {
    const factory = new IDBFactory();
    const persist = await openLocalPersistence(factory);
    await persist.putLinks([link]);
    await persist.putInbox([row]);
    await persist.deleteLinks([link.id]);
    expect((await persist.load()).links).toEqual([]);
    await persist.clear();
    expect((await persist.load()).inbox).toEqual([]);
  });

  it("falls back to memory without IndexedDB or when opening fails", async () => {
    expect((await openLocalPersistence(undefined)).durable).toBe(false);
    const refusing = { open: () => { throw new Error("SecurityError"); } } as unknown as IDBFactory;
    expect((await openLocalPersistence(refusing)).durable).toBe(false);
  });
});
```

- [ ] **Step 3: Run to verify they fail**

Run: `npx vitest run src/link/local-persist.test.ts`
Expected: FAIL, `openLocalPersistence` is not exported.

- [ ] **Step 4: Implement (append to `src/link/local-persist.ts`)**

```ts
const DB_NAME = "plsfix-links";
const DB_VERSION = 1;
const LINKS = "links";
const INBOX = "inbox";

export async function openLocalPersistence(
  factory: IDBFactory | undefined = globalThis.indexedDB,
): Promise<LocalPersistence> {
  if (factory === undefined) return memoryPersistence();
  try {
    return new IdbPersistence(await openDb(factory));
  } catch {
    return memoryPersistence();
  }
}

function openDb(factory: IDBFactory): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = factory.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(LINKS)) db.createObjectStore(LINKS, { keyPath: "id" });
      if (!db.objectStoreNames.contains(INBOX)) db.createObjectStore(INBOX, { keyPath: "id" });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("indexedDB: open failed"));
    request.onblocked = () => reject(new Error("indexedDB: open blocked"));
  });
}

function readAll<T>(db: IDBDatabase, name: string): Promise<T[]> {
  return new Promise((resolve, reject) => {
    const request = db.transaction(name, "readonly").objectStore(name).getAll();
    request.onsuccess = () => resolve(request.result as T[]);
    request.onerror = () => reject(request.error ?? new Error(`indexedDB: read ${name}`));
  });
}

class IdbPersistence implements LocalPersistence {
  durable = true;

  constructor(private readonly db: IDBDatabase) {}

  async load(): Promise<{ links: StoredLocalLink[]; inbox: StoredLocalInbox[] }> {
    try {
      const [links, inbox] = await Promise.all([
        readAll<StoredLocalLink>(this.db, LINKS),
        readAll<StoredLocalInbox>(this.db, INBOX),
      ]);
      return { links, inbox };
    } catch {
      this.durable = false;
      return { links: [], inbox: [] };
    }
  }

  putLinks(links: StoredLocalLink[]): Promise<void> {
    return this.write([LINKS], (tx) => links.forEach((link) => tx.objectStore(LINKS).put(link)));
  }
  putInbox(rows: StoredLocalInbox[]): Promise<void> {
    return this.write([INBOX], (tx) => rows.forEach((row) => tx.objectStore(INBOX).put(row)));
  }
  deleteLinks(ids: string[]): Promise<void> {
    return this.write([LINKS], (tx) => ids.forEach((id) => tx.objectStore(LINKS).delete(id)));
  }
  deleteInbox(ids: string[]): Promise<void> {
    return this.write([INBOX], (tx) => ids.forEach((id) => tx.objectStore(INBOX).delete(id)));
  }
  clear(): Promise<void> {
    return this.write([LINKS, INBOX], (tx) => {
      tx.objectStore(LINKS).clear();
      tx.objectStore(INBOX).clear();
    });
  }

  // Resolves either way: a full or refused store must not fail the paste that
  // is already in memory; it only stops claiming to be durable.
  private write(names: string[], fill: (tx: IDBTransaction) => void): Promise<void> {
    return new Promise((resolve) => {
      try {
        const tx = this.db.transaction(names, "readwrite");
        tx.oncomplete = () => resolve();
        tx.onerror = tx.onabort = () => {
          this.durable = false;
          resolve();
        };
        fill(tx);
      } catch {
        this.durable = false;
        resolve();
      }
    });
  }
}
```

- [ ] **Step 5: Run to verify they pass**

Run: `npx vitest run src/link/local-persist.test.ts`
Expected: PASS (3 tests). If `fake-indexeddb` returns plain objects where a `Uint8Array` went in, compare with
`Array.from(...)` in the test and say so in the report; do not change the adapter to arrays.

- [ ] **Step 6: Commit**

```bash
git add package.json package-lock.json src/link/local-persist.ts src/link/local-persist.test.ts
git commit -m "IndexedDB persistence for pasted links" -m "src/link/local-persist.ts: database plsfix-links (links, inbox), write-through, never throws into a paste; memory when the webview has no IndexedDB or refuses it. fake-indexeddb for the tests."
```

### Task 7: Revert in either rev space, and two neutral sentences

**Files:**
- Modify: `src/ppt/revert.ts` (`previousRev` uses `previousRevOf`; the "gone" sentence)
- Modify: `src/link/status.ts` (`StaleRelayError` message)
- Test: the existing revert and status tests (grep `no longer holds version` and `older picture`)

- [ ] **Step 1: Update the tests to the new sentences and add the local case**

The two existing assertions change (report them as wording changes):
- `"revert <label>: the relay no longer holds version N."` -> `"revert <label>: the previous version is no longer available."`
- `"The relay sent an older picture than this deck already holds."` -> `"The update is older than the picture this deck already holds."`

Add to the revert test file:

```ts
it("counts the first local revision as having no previous version", async () => {
  // Build one row whose tag.rev is LOCAL_REV_BASE + 1 with the helpers this file
  // already uses, run revertLinks against any relay, and expect
  // { reverted: 0, noPrevious: 1, failed: 0 }.
});
```

Write that test with the file's own row builder and relay stub; the comment above is the
behaviour, not a placeholder to leave in.

- [ ] **Step 2: Run to verify they fail**

Run: `npx vitest run src/ppt src/link/status.test.ts`
Expected: FAIL on the three changed or new assertions.

- [ ] **Step 3: Implement**

In `src/ppt/revert.ts`:

```ts
import { previousRevOf } from "../link/local";
// ...
// The relay (or this computer's store) holds the tag's revision and the one
// before it, so a link that was never updated - or one already reverted - has
// nowhere left to go. Local revs stop at the first one of their own space.
function previousRev(found: FoundLink): number | null {
  return previousRevOf(found.tag.rev);
}
```

and in `countFailure` the gone line becomes
`` `revert ${sourceLabel(found.tag.src, found.tag.kind)}: the previous version is no longer available.` ``
(the `rev` parameter stays for the signature; mark it `_rev` if lint asks).

In `src/link/status.ts`: `super("The update is older than the picture this deck already holds.");`

- [ ] **Step 4: Run to verify they pass**

Run: `npx vitest run src/ppt src/link test/ppt*.ts test/stress.ppt*.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/ppt/revert.ts src/link/status.ts src/ppt src/link test
git commit -m "Revert in either rev space; two sentences that fit both transports" -m "src/ppt/revert.ts uses previousRevOf (local revs stop at their first); the gone sentence and StaleRelayError no longer name the relay. Wording changes listed in the slice report."
```

### Task 8: The two-flavor clipboard write and the paste read

**Files:**
- Create: `src/ui/clipboard-links.ts`
- Test: `src/ui/clipboard-links.test.ts` (`// @vitest-environment jsdom`)

**Interfaces:**
- Consumes: `bundleHtml`, `BUNDLE_SENTENCE`, `readPastedBundle`, `BundleRead` (Task 1).
- Produces: `interface PendingCopy { resolve(json: string): void; reject(error: unknown): void; done: Promise<boolean> }`,
  `beginClipboardWrite(): PendingCopy`, `copyBundleNow(json: string): boolean`,
  `bundleFromPaste(data: DataTransfer | null): BundleRead`.

- [ ] **Step 1: Write the failing tests**

```ts
// @vitest-environment jsdom
// src/ui/clipboard-links.test.ts
// The write starts inside the press and its content arrives later; the
// synchronous fallback sets both flavors; a paste reads the bundle back.
import { afterEach, describe, expect, it, vi } from "vitest";
import { BUNDLE_SENTENCE, bundleHtml } from "../link/bundle";
import { beginClipboardWrite, bundleFromPaste, copyBundleNow } from "./clipboard-links";

class FakeItem {
  constructor(readonly items: Record<string, Promise<Blob>>) {}
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("beginClipboardWrite", () => {
  it("calls write at once and lands both flavors when the content arrives", async () => {
    const write = vi.fn(async (items: FakeItem[]) => {
      const blobs = await Promise.all(Object.values(items[0]!.items));
      expect(await blobs[0]!.text()).toBe(bundleHtml("{}"));
      expect(await blobs[1]!.text()).toBe(BUNDLE_SENTENCE);
    });
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", { clipboard: { write } });
    const pending = beginClipboardWrite();
    expect(write).toHaveBeenCalledTimes(1);
    pending.resolve("{}");
    await expect(pending.done).resolves.toBe(true);
  });

  it("answers false when the engine has no ClipboardItem or refuses", async () => {
    vi.stubGlobal("ClipboardItem", undefined);
    await expect(beginClipboardWrite().done).resolves.toBe(false);
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", { clipboard: { write: async () => { throw new DOMException("no", "NotAllowedError"); } } });
    const pending = beginClipboardWrite();
    pending.resolve("{}");
    await expect(pending.done).resolves.toBe(false);
  });
});

describe("copyBundleNow", () => {
  it("sets both flavors through a copy event", () => {
    const data = new Map<string, string>();
    document.execCommand = vi.fn(() => {
      const event = new Event("copy", { cancelable: true });
      Object.defineProperty(event, "clipboardData", { value: { setData: (t: string, v: string) => data.set(t, v) } });
      document.dispatchEvent(event);
      return true;
    });
    expect(copyBundleNow("{}")).toBe(true);
    expect(data.get("text/html")).toBe(bundleHtml("{}"));
    expect(data.get("text/plain")).toBe(BUNDLE_SENTENCE);
  });
});

describe("bundleFromPaste", () => {
  it("reads nothing from nothing", () => {
    expect(bundleFromPaste(null)).toEqual({ ok: false, reason: "notBundle" });
  });
});
```

- [ ] **Step 2: Run to verify they fail**

Run: `npx vitest run src/ui/clipboard-links.test.ts`
Expected: FAIL, unresolved import.

- [ ] **Step 3: Implement**

```ts
// src/ui/clipboard-links.ts
// The clipboard for link bundles: text/html carries the bundle in one data
// attribute, text/plain only the sentence, so a paste anywhere but the
// pls,fix pane gives one line. Invariant: the write starts inside the press
// (WebKit requires it); the content may arrive after the Excel work.

import {
  BUNDLE_SENTENCE,
  bundleHtml,
  readPastedBundle,
  type BundleRead,
} from "../link/bundle";

export interface PendingCopy {
  resolve(json: string): void;
  reject(error: unknown): void;
  // True once the write landed; false when the engine has no ClipboardItem
  // or refused it (focus moved, a permission): the caller then offers the
  // synchronous copy behind a second press.
  done: Promise<boolean>;
}

export function beginClipboardWrite(): PendingCopy {
  let resolveJson!: (json: string) => void;
  let rejectJson!: (error: unknown) => void;
  const json = new Promise<string>((resolve, reject) => {
    resolveJson = resolve;
    rejectJson = reject;
  });
  json.catch(() => undefined);
  const Item = (globalThis as { ClipboardItem?: typeof ClipboardItem }).ClipboardItem;
  const clipboard = (globalThis as { navigator?: Navigator }).navigator?.clipboard;
  let done: Promise<boolean> = Promise.resolve(false);
  if (Item !== undefined && clipboard?.write !== undefined) {
    try {
      const html = json.then((text) => new Blob([bundleHtml(text)], { type: "text/html" }));
      const plain = json.then(() => new Blob([BUNDLE_SENTENCE], { type: "text/plain" }));
      done = clipboard
        .write([new Item({ "text/html": html, "text/plain": plain })])
        .then(() => true, () => false);
    } catch {
      done = Promise.resolve(false);
    }
  }
  return { resolve: resolveJson, reject: rejectJson, done };
}

// Inside a press only: execCommand("copy") is refused anywhere else.
export function copyBundleNow(json: string): boolean {
  let wrote = false;
  const onCopy = (event: Event): void => {
    const data = (event as ClipboardEvent).clipboardData;
    if (!data) return;
    data.setData("text/html", bundleHtml(json));
    data.setData("text/plain", BUNDLE_SENTENCE);
    event.preventDefault();
    wrote = true;
  };
  document.addEventListener("copy", onCopy);
  try {
    return document.execCommand("copy") && wrote;
  } catch {
    return false;
  } finally {
    document.removeEventListener("copy", onCopy);
  }
}

// Read inside the paste event: its DataTransfer is empty once the event ends.
export function bundleFromPaste(data: DataTransfer | null): BundleRead {
  if (data === null) return { ok: false, reason: "notBundle" };
  return readPastedBundle(data.getData("text/html"), data.getData("text/plain"));
}
```

- [ ] **Step 4: Run to verify they pass**

Run: `npx vitest run src/ui/clipboard-links.test.ts`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add src/ui/clipboard-links.ts src/ui/clipboard-links.test.ts
git commit -m "Two-flavor clipboard for link bundles" -m "src/ui/clipboard-links.ts: a write started inside the press with the content arriving later, a synchronous copy-event fallback, and the paste read (HTML flavor first). Plain text is only the sentence."
```

### Task 9: The per-device transport setting

**Files:**
- Create: `src/link/transport-setting.ts`
- Test: `src/link/transport-setting.test.ts`

**Interfaces:**
- Consumes: `KeyStore`, `WORKSPACE_STORAGE_KEY` (`src/link/workspace.ts`).
- Produces: `type LinkTransport = "local" | "relay"`, `TRANSPORT_STORAGE_KEY`,
  `loadTransport(store: KeyStore): Promise<LinkTransport>`, `saveTransport(store: KeyStore, transport: LinkTransport): Promise<void>`.

- [ ] **Step 1: Write the failing tests**

```ts
// src/link/transport-setting.test.ts
// Local by default; a device that already holds a link key keeps the relay.
import { describe, expect, it } from "vitest";
import { loadTransport, saveTransport, TRANSPORT_STORAGE_KEY } from "./transport-setting";
import { WORKSPACE_STORAGE_KEY, type KeyStore } from "./workspace";

function store(seed: Record<string, string> = {}): KeyStore {
  const map = new Map(Object.entries(seed));
  return {
    get: async (key) => map.get(key) ?? null,
    set: async (key, value) => void map.set(key, value),
    remove: async (key) => void map.delete(key),
  };
}

describe("loadTransport", () => {
  it("is local on a new install", async () => {
    expect(await loadTransport(store())).toBe("local");
  });
  it("keeps the relay where a link key is already stored", async () => {
    expect(await loadTransport(store({ [WORKSPACE_STORAGE_KEY]: "k" }))).toBe("relay");
  });
  it("follows what was saved, whatever the key says", async () => {
    const keys = store({ [WORKSPACE_STORAGE_KEY]: "k" });
    await saveTransport(keys, "local");
    expect(await keys.get(TRANSPORT_STORAGE_KEY)).toBe("local");
    expect(await loadTransport(keys)).toBe("local");
  });
  it("treats an unreadable store as a new install", async () => {
    const broken: KeyStore = { get: async () => { throw new Error("denied"); }, set: async () => undefined, remove: async () => undefined };
    expect(await loadTransport(broken)).toBe("local");
  });
});
```

- [ ] **Step 2: Run to verify they fail**

Run: `npx vitest run src/link/transport-setting.test.ts`
Expected: FAIL, unresolved import.

- [ ] **Step 3: Implement**

```ts
// src/link/transport-setting.ts
// Which way links travel on this device: "local" (one copy, one paste) or
// "relay". Stored per device beside the link key. Invariant: nothing stored
// means relay only where a key is already here, so existing users keep what
// works and new installs start local.

import { WORKSPACE_STORAGE_KEY, type KeyStore } from "./workspace";

export type LinkTransport = "local" | "relay";

export const TRANSPORT_STORAGE_KEY = "plsfix.link.transport.v1";

async function read(store: KeyStore, key: string): Promise<string | null> {
  try {
    return await store.get(key);
  } catch {
    return null;
  }
}

export async function loadTransport(store: KeyStore): Promise<LinkTransport> {
  const saved = await read(store, TRANSPORT_STORAGE_KEY);
  if (saved === "local" || saved === "relay") return saved;
  return (await read(store, WORKSPACE_STORAGE_KEY)) === null ? "local" : "relay";
}

export async function saveTransport(store: KeyStore, transport: LinkTransport): Promise<void> {
  await store.set(TRANSPORT_STORAGE_KEY, transport);
}
```

- [ ] **Step 4: Run to verify they pass**

Run: `npx vitest run src/link/transport-setting.test.ts`
Expected: PASS (4 tests).

- [ ] **Step 5: Slice gates, then commit**

Run: `npm run check && npm run ux:sweep -- --port 3172 && npm run ux:check -- --port 3171`
Expected: all green.

```bash
git add src/link/transport-setting.ts src/link/transport-setting.test.ts
git commit -m "Per-device link transport setting" -m "src/link/transport-setting.ts: local or relay, stored under plsfix.link.transport.v1; nothing stored = relay where a link key exists, local otherwise."
```

---

## Slice S2: the Excel pane (worktree `local-links-excel`, ports 3181/3182)

### Task 10: Copies carry an inbox row per link

**Files:**
- Modify: `src/excel/links.ts` (`pushLinks`, `pushRegistry`, `pushOne`)
- Modify: `src/excel/link-record.ts` (export `announce`)
- Test: `test/links.local.integration.test.ts` (new; strict fake host)

**Interfaces:**
- Consumes: `LocalCollector` (Task 3), `localWorkspace` (Task 2), `decodeBundle` (Task 1).
- Produces: `pushLinks(ids: string[] | "all", relay: RelayApi, options?: { announce?: Workspace }): Promise<PushSummary>`
  (same for `pushRegistry`). With `announce`, every pushed link also posts its inbox row sealed with that workspace.

- [ ] **Step 1: Write the failing test**

A new integration file over the strict fake host (copy the boot pattern from `test/links.integration.test.ts`):
1. Seed a small block, select it, `await exportSelection(await localWorkspace(), collector)` with a `new LocalCollector()`;
   expect `collector.bundle()` to hold 1 link with `rev === LOCAL_REV_BASE + 1` and 1 inbox row; open the row
   with `open(ws.enc, ws.id, row.blob)` + `decodeInboxItem` and open the link blob with the item's token
   (`deriveLinkKeys(item.token)`, `open(keys.enc, id, blob)`, `decodePayload`) and expect a picture payload.
2. Change a cell, `await pushLinks("all", second, { announce: ws })` with a fresh collector; expect 1 link at
   `LOCAL_REV_BASE + 2` and 1 inbox row, and the workbook registry's `rev` equal to it.
3. `await pushLinks("all", third)` without `announce`; expect links and no inbox rows (relay behaviour kept).

- [ ] **Step 2: Run to verify it fails**

Run: `npx vitest run test/links.local.integration.test.ts`
Expected: FAIL at step 2 (no inbox rows: `announce` does not exist yet).

- [ ] **Step 3: Implement**

`src/excel/link-record.ts`: change `async function announce(` to `export async function announce(`.

`src/excel/links.ts`:

```ts
export interface PushOptions {
  // Local mode: also post each pushed link's inbox row, sealed with this
  // workspace, so a deck that does not hold the link yet can insert it.
  announce?: Workspace;
}

export async function pushLinks(
  ids: string[] | "all",
  relay: RelayApi,
  options: PushOptions = {},
): Promise<PushSummary> {
  return exclusive("push", () => pushRegistry(ids, relay, options));
}
```

Thread `options` through `pushRegistry(ids, relay, options)` into `pushOne(..., relay, summary, options)`,
and after `entry.lastPushedAt = new Date().toISOString();` in `pushOne` add:

```ts
    if (options.announce) await announce(entry, src, options.announce, relay);
```

Auto-push (`src/excel/link-watch.ts`) keeps calling `pushRegistry(ids, relay)`: no change.

- [ ] **Step 4: Run to verify it passes**

Run: `npx vitest run test/links.local.integration.test.ts test/links*.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/excel/links.ts src/excel/link-record.ts test/links.local.integration.test.ts
git commit -m "Copy selected / Copy all can carry an inbox row per link" -m "src/excel/links.ts: pushLinks takes { announce } and posts each pushed link's inbox row with it; link-record.ts exports announce. Relay pushes unchanged."
```

### Task 11: The Excel transport module and the local send paths

**Files:**
- Create: `src/pane/links-transport.ts`
- Modify: `src/pane/links-tab.ts` (wiring only: `exportRange`, `exportChart`, `push`, the setting)
- Modify: `taskpane.html` (Links tab: the bar in the Export section, the new last section)
- Modify: the stylesheet that owns the Links tab (find it with `grep -rn "link-key\|key-display" src/styles`)
- Modify: `src/help/copy.excel.ts` (one sentence for `copy-for-powerpoint`)
- Test: `src/pane/links-transport.test.ts` (jsdom), `test/clickthrough.excel.integration.test.ts` (coverage)

**Interfaces:**
- Consumes: `beginClipboardWrite`, `copyBundleNow` (Task 8), `LocalCollector` (Task 3), `localWorkspace` (Task 2),
  `encodeBundle`, `BUNDLE_MAX_CHARS` (Task 1), `loadTransport`, `saveTransport`, `LinkTransport` (Task 9),
  `pushLinks(..., { announce })` (Task 10).
- Produces (used by `links-tab.ts` only):

```ts
export interface LinksTransport {
  mode(): LinkTransport;
  setMode(mode: LinkTransport): Promise<void>;
  // Local mode only. Must be called synchronously inside the press: it starts
  // the clipboard write before its first await, then runs the Excel work
  // against a LocalCollector and the local workspace.
  copy<T>(
    run: (ws: Workspace, relay: RelayApi) => Promise<T>,
    lines: { copied: (result: T) => string; ready: (result: T) => string },
  ): Promise<string>;
}
export function installLinksTransport(deps: { root: ParentNode; keyStore: KeyStore; toast: Toast }): Promise<LinksTransport>;
```

Markup (the `copy-ready` bar goes at the end of the Export section, the new section after the Linked
objects section and before the Link key section; give the Link key `<section>` the id `link-key-section`
and the Auto-push label the id `links-autopush-row`):

```html
<div id="copy-ready" class="copy-ready" hidden>
  <p id="copy-ready-text" class="hint"></p>
  <button id="copy-for-powerpoint" type="button">Copy for PowerPoint</button>
  <textarea id="copy-manual" class="copy-manual" readonly hidden aria-label="The copy for PowerPoint, to select and copy by hand"></textarea>
</div>
```

```html
<section aria-labelledby="transport-heading" data-help="link-transport">
  <div class="section-heading compact">
    <div>
      <p class="kicker">HOW LINKS REACH POWERPOINT</p>
      <h2 id="transport-heading">Transport</h2>
    </div>
  </div>
  <label class="key-label" for="link-transport">Links travel
    <select id="link-transport">
      <option value="local">Copy and paste on this computer</option>
      <option value="relay">Through the relay</option>
    </select>
  </label>
  <p class="hint" id="transport-local-hint">Nothing leaves this computer: export or copy here, then paste in PowerPoint's pls,fix pane, Inbox tab.</p>
  <p class="hint" id="transport-relay-hint" hidden>Pushes go to the relay beside this pane; PowerPoint updates with one click once the link key is pasted there.</p>
</section>
```

`copy()` behaviour, in order:
1. `const pending = beginClipboardWrite();` (first statement, no await before it).
2. `const collector = new LocalCollector();` then `const result = await run(await localWorkspace(), collector);`
3. `const json = encodeBundle(collector.bundle());` if `json.length > BUNDLE_MAX_CHARS` then
   `pending.reject(...)` and throw `"That is too much to copy at once: copy one project or the selected links."`.
   If `collector.isEmpty()` then `pending.reject(...)` and return `"Nothing to copy: no link could be read."`.
4. `pending.resolve(json)`; `if (await pending.done) { hide the bar; return lines.copied(result); }`
5. Otherwise keep `json` as the prepared copy, show the bar with
   `copy-ready-text` = `lines.ready(result)` and return `lines.ready(result)`.
6. Any throw from `run` or step 3: `pending.reject(error)`, rethrow (the guard reports it).

The `copy-for-powerpoint` press (through the tab's `wire()`, so the busy latch covers it):
`copyBundleNow(prepared)` synchronously; true -> hide the bar, `"Copied for PowerPoint. Paste it in the pls,fix pane, Inbox tab."`;
false -> show `copy-manual` with the JSON, select it, throw `"Copy failed: select the text below and copy it by hand (Ctrl+C, or ⌘C on a Mac)."`.

`links-tab.ts` call sites (the only logic that changes there):

```ts
async function exportRange(tab: Tab, kind: ExportKind): Promise<string> {
  if (tab.transport.mode() === "local") {
    return tab.transport.copy((ws, relay) => sender(kind)(ws, relay), localExportLines(tab));
  }
  // relay path unchanged
}
```

with `localExportLines` returning
`copied: (r) => \`Copied for PowerPoint: ${r.label}. Paste it in PowerPoint: pls,fix, Inbox tab.\``
and `ready: (r) => \`${r.label} is linked and ready: press Copy for PowerPoint.\``, each followed by
`await refresh(tab)` inside `run`. `exportChart` the same. `push(tab, action, all)` in local mode runs
`tab.transport.copy((ws, relay) => pushLinks(ids, relay, { announce: ws }), ...)` with the summary words
"copied" instead of "pushed" (`summarize` gets a `verb` argument). Compute `ids` (which may throw
"Select a link in the list first.") BEFORE calling `copy()`.

- [ ] **Step 1: Write the failing tests** (`src/pane/links-transport.test.ts`, jsdom): with `navigator.clipboard.write`
  stubbed to succeed, `copy()` calls `write` before its first await and returns the copied line; stubbed to refuse,
  it shows `#copy-ready` with the ready line and `copy-for-powerpoint` then copies through a stubbed
  `execCommand` copy event; with `execCommand` failing too, `#copy-manual` shows the JSON; a `run` that throws
  rejects the pending write and rethrows. Seed `installLinksTransport` with an in-memory `KeyStore`.
- [ ] **Step 2: Run to verify they fail.** `npx vitest run src/pane/links-transport.test.ts` -> FAIL.
- [ ] **Step 3: Implement** `src/pane/links-transport.ts` (header comment; under 200 lines) and the `links-tab.ts`
  call sites; add the markup and the one help line
  `"copy-for-powerpoint": "Copies what the last export or copy prepared, when the first try could not reach the clipboard."`
  plus a `link-transport` section entry in `src/help/copy.excel.ts` (`about` only, no buttons).
- [ ] **Step 4: Update `test/clickthrough.excel.integration.test.ts`** so `copy-for-powerpoint` is pressed (it is a
  static button; in state (a) its press answers the "nothing prepared" sentence
  `"Nothing is waiting to be copied."`, which `copy-for-powerpoint` throws when no bundle is prepared).
- [ ] **Step 5: Run** `npx vitest run src/pane test/clickthrough.excel.integration.test.ts test/I.audit.integration.test.ts` -> PASS.
- [ ] **Step 6: Commit** "Excel local mode: export and copy through the clipboard" (body: files + behaviour).

### Task 12: The setting, the labels and what hides in each mode

**Files:**
- Modify: `src/pane/links-transport.ts` (`applyMode(mode)`: visibility and labels), `src/pane/links-tab.ts` (the select's change handler, wiring only)
- Test: `src/pane/links-transport.test.ts`, `test/clickthrough.excel.integration.test.ts`

Behaviour of `applyMode`:
- `local`: `#link-key-section` hidden, `#links-autopush-row` hidden (and auto-push switched off through the
  existing toggle path if it was on, with the toast "Auto-push needs the relay: it is off in copy and paste mode."),
  `#push-selected` text "Copy selected", `#push-all` text "Copy all", `#transport-local-hint` shown,
  `#transport-relay-hint` hidden, the select shows `local`.
- `relay`: the reverse; labels "Push selected" / "Push all"; the bar `#copy-ready` hidden.
- The select's `change` runs through `guarded(tab, "link-transport", ...)`: `saveTransport`, `applyMode`,
  toast `"Links now travel by copy and paste on this computer."` or `"Links now travel through the relay."`.
- Boot: `installLinksTransport` reads `loadTransport` and applies it before the tab's first render.

- [ ] **Step 1: Write the failing tests**: each mode's visibility and labels; the change handler saves and toasts;
  a device with a stored key boots in relay mode, one without in local mode (boot the pane twice over the fake
  host with and without `plsfix.link.workspace.v1` in `localStorage`, the pattern in `test/I.audit.integration.test.ts`).
- [ ] **Step 2: Run to verify they fail.**
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run** the pane tests and both click-through suites -> PASS; then the slice gates:
  `npm run check && npm run ux:sweep -- --port 3182 && npm run ux:check -- --port 3181`.
- [ ] **Step 5: Commit** "Excel local mode: the transport select, labels and visibility".

---

## Slice S3: the PowerPoint pane (worktree `local-links-ppt`, ports 3191/3192)

### Task 13: The PowerPoint transport and the Settings select

**Files:**
- Create: `src/ppt/transport.ts`
- Modify: `src/ppt/main.ts` (wiring only: every `relay` use goes through `transport.relay()`, `requireWorkspace()` through `transport.workspace()`)
- Modify: `pptpane.html` (Settings: the select above the key controls; id `link-key-controls` on the wrapper of the key input, buttons and hints)
- Test: `src/ppt/transport.test.ts`, `test/stress.ppt.pane.integration.test.ts` (both modes)

**Interfaces:**
- Consumes: `LocalStore`, `openLocalPersistence` (Tasks 5-6), `localWorkspace` (Task 2), `loadTransport`, `saveTransport` (Task 9).
- Produces:

```ts
export interface PptTransport {
  mode(): LinkTransport;
  relay(): RelayApi;               // the RelayClient, or this computer's LocalStore
  workspace(): Workspace | null;   // the paired key, or the local workspace
  store(): LocalStore | null;      // local mode's store, once open
  setMode(mode: LinkTransport): Promise<void>;
}
export function openPptTransport(deps: {
  keyStore: KeyStore;
  remote: RelayApi;
  paired: () => Workspace | null;
  persistence?: () => Promise<LocalPersistence>;   // defaults to openLocalPersistence
}): Promise<PptTransport>;
```

Rules: `openPptTransport` loads the mode; the LocalStore opens lazily the first time local mode is used
and is kept. In `main.ts`, `const relay = new RelayClient(...)` becomes `const remote = new RelayClient(...)`;
`openPptTransport` runs in `Office.onReady` BEFORE `ready = true` (so no button reaches a transport that is not
there); every former `relay` argument becomes `transport.relay()`; `requireWorkspace()` returns
`transport.workspace()` or throws `PAIR_FIRST` (relay mode) as before; `renderPairing()` in local mode says
`"Copy and paste on this computer: no key needed."`; `renderInboxView()` shows the list whenever
`transport.workspace() !== null`. The Settings select change: `setMode`, re-render, `reloadLinks()`,
`inboxQuietly()`, toast `"Links now travel by copy and paste on this computer."` / `"Links now travel through the relay."`.
In local mode `#link-key-controls` is hidden.

- [ ] **Step 1: Write the failing tests**: `transport.test.ts` for the mode rules and the lazy store (a persistence
  stub counts opens); `stress.ppt.pane` boots once per mode (local: no key stored; relay: key stored) and
  presses `refresh-links` and `refresh-inbox`, expecting no network call in local mode (the FakeRelay's call log
  stays empty) and "Nothing waiting from Excel." from the local store.
- [ ] **Step 2: Run to verify they fail.**
- [ ] **Step 3: Implement** `src/ppt/transport.ts` (header comment; under 150 lines) and the wiring.
- [ ] **Step 4: Run** `npx vitest run src/ppt test/ppt*.ts test/stress.ppt*.ts test/clickthrough.ppt.integration.test.ts` -> PASS.
- [ ] **Step 5: Commit** "PowerPoint local mode: the transport and the Settings select".

### Task 14: Paste from Excel

**Files:**
- Create: `src/ppt/paste-links.ts`
- Test: `test/ppt.paste-links.integration.test.ts` (fakeppt + LocalStore + memory persistence)

**Interfaces:**
- Consumes: `LocalStore.ingest` (Task 5), `listLinks`, `updateLinks` (`src/ppt/links.ts`), `PptHost`.
- Produces:

```ts
export interface PasteSummary {
  links: number;       // links the bundle carried
  updated: number;
  current: number;
  waiting: number;     // new inbox rows kept for links the deck does not hold
  failed: number;
  failures: string[];
  notes: string[];
}
export async function pasteLinks(bundle: Bundle, store: LocalStore, host?: PptHost): Promise<PasteSummary>;
export function summarizePaste(summary: PasteSummary): string;
```

Flow of `pasteLinks`: `const found = await host.scanLinks();` -> `deckIds` from `found` -> `store.ingest(bundle, deckIds)`
-> `rows = await listLinks(store, once)` where `once` is `Object.assign(Object.create(host), { scanLinks: async () => found })`
(one scan per paste) -> `subset = rows.filter((row) => ids.has(row.found.tag.id))` -> `updateLinks(subset, store, host)`
-> map the `UpdateSummary` into `PasteSummary`.

`summarizePaste`: `"Pasted N links: A updated, B up to date, C waiting in the Inbox"` with zero parts left out,
`", D failed"` when failures, and a final full stop; `N` counts `bundle.links.length` ("1 link" singular).
With no links but inbox rows: `"Pasted: C waiting in the Inbox."`.

- [ ] **Step 1: Write the failing tests**: (a) a bundle whose link the deck holds at an older rev repaints it
  (tag rev becomes the bundle's rev) and says "1 updated"; (b) the same bundle pasted twice says "1 up to date";
  (c) a bundle with a new link puts it in the Inbox ("1 waiting in the Inbox") and `insertFromInbox(item, localWs, store)`
  inserts it; (d) an older bundle after a newer one downgrades nothing; (e) a blob that will not open fails its
  row only, with a failure line. Build bundles with `LocalCollector` + the Excel-side sealing helpers
  (`seal`, `deriveLinkKeys`, `encodePayload`, `encodeInboxItem`), the way `test/stress.ppt.support.ts` builds exports.
- [ ] **Step 2: Run to verify they fail.**
- [ ] **Step 3: Implement** (header comment; under 150 lines; no Office.js outside `host`).
- [ ] **Step 4: Run** -> PASS.
- [ ] **Step 5: Commit** "PowerPoint: paste a bundle from Excel".

### Task 15: The paste box, Clear pasted links, the "Sent" column

**Files:**
- Modify: `pptpane.html` (Inbox tab: the paste box and Clear pasted links before the Slide picker; Links table header "Pushed" -> "Sent")
- Modify: `src/ppt/main.ts` (wiring: the paste listener, the new button in `BUTTON_ACTIONS` + `CONFIRM_BUTTON_IDS`, local-mode visibility)
- Modify: `src/styles/ppt.css` (the paste box: dashed border from an existing token, min-height 56px, full width)
- Modify: `src/help/copy.ppt.ts` (one sentence for `clear-pasted-links`; the `update-*` and `refresh-links` sentences stay)
- Test: `test/clickthrough.ppt.integration.test.ts`, `test/stress.ppt.pane.integration.test.ts`

Markup:

```html
<div id="paste-links-area" hidden>
  <label class="paste-label" for="paste-links">Paste from Excel</label>
  <textarea id="paste-links" class="paste-box" rows="2" spellcheck="false"
    placeholder="Click here, then Ctrl+V (⌘V on a Mac)"></textarea>
  <button id="clear-pasted-links" type="button">Clear pasted links</button>
  <p class="hint" id="paste-not-durable" hidden>This computer does not keep pasted links between sessions: paste again after reopening the pane.</p>
</div>
```

Wiring in `main.ts` (wiring only; the logic is `pasteLinks`):

```ts
getElement<HTMLTextAreaElement>("paste-links").addEventListener("paste", (event) => {
  event.preventDefault();
  const read = bundleFromPaste(event.clipboardData);   // inside the event: the data is gone after it
  act(async () => {
    if (!read.ok) {
      throw new Error(read.reason === "newerVersion"
        ? "This copy comes from a newer pls,fix. Update the add-in."
        : "That is not a pls,fix copy from Excel.");
    }
    const store = transport.store();
    if (store === null) throw new Error("Switch Settings to copy and paste first.");
    const summary = await pasteLinks(read.bundle, store);
    await refreshQuietly();
    await inboxQuietly();
    for (const line of summary.failures) details.addLine(line);   // use the details API the file already uses
    return summarizePaste(summary);
  }, "paste-links");
});
getElement<HTMLTextAreaElement>("paste-links").addEventListener("input", (event) => {
  (event.target as HTMLTextAreaElement).value = "";   // typing never fills the box
});
```

`clear-pasted-links` (in `BUTTON_ACTIONS` and `CONFIRM_BUTTON_IDS`): `await transport.store()?.clear()`, empty
`inboxItems`, re-render, `await refreshQuietly()`, return `"Pasted links cleared from this computer. The deck keeps its objects."`.
`#paste-links-area` shows only in local mode; `#paste-not-durable` shows when `!store.durable()`.
Help line: `"clear-pasted-links": "Deletes every pasted link and waiting item kept on this computer. The deck keeps its objects. Press twice."`

- [ ] **Step 1: Write the failing tests**: in the PowerPoint click-through suite, a local-mode state where a
  synthetic `paste` event (a `DataTransfer`-like object with `getData`) carrying a bundle built as in Task 14
  lands in the Inbox and updates the deck; a non-bundle paste toasts "That is not a pls,fix copy from Excel.";
  `clear-pasted-links` pressed twice empties the Inbox; coverage includes the new button.
- [ ] **Step 2: Run to verify they fail.**
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Slice gates**: `npm run check && npm run ux:sweep -- --port 3192 && npm run ux:check -- --port 3191`.
- [ ] **Step 5: Commit** "PowerPoint local mode: the paste box and Clear pasted links".

---

## Slice S4: words and permanent proofs (on main after S2 and S3 merge)

### Task 16: Help, README, INSTALL, privacy, manual, Map

- [ ] Help copy (`src/help/copy.excel.ts`, `src/help/copy.ppt.ts`): every sentence that says "relay", "push" or
  "link key" is true in both modes or names its mode ("In relay mode, ..."); sections `link-transport` (Excel)
  and the PowerPoint Settings about line explain the two modes in one sentence each. Run `npx vitest run src/help`.
- [ ] `README.md` (Daniel's hand-curated file: surgical edits only): one short paragraph under the links
  section: local by default, relay optional, where the switch is.
- [ ] `docs/INSTALL.md`: one line that a fresh install needs no relay for links.
- [ ] `public/privacy.html`: a paragraph on local mode (nothing sent; pasted links kept in this browser's
  IndexedDB until Clear pasted links; a copied bundle is private).
- [ ] `manual/src/content/links.rs` (Latvian, normal register): a section "Saites bez servera: kopēt un ielīmēt"
  with the steps (Excel: Export vai Copy all; PowerPoint: pls,fix, Inbox, ielīmēšanas lauks, Ctrl+V) and the
  switch; `cargo test --manifest-path manual/Cargo.toml`.
- [ ] `CLAUDE.md` Map: the new modules (one line each in their module lines), the data flow for local mode, the
  symptom lines ("a paste does nothing" -> `bundleFromPaste` + the paste box; "a copy never lands" ->
  `beginClipboardWrite` / the Copy for PowerPoint bar; "Not pasted yet on every row" -> the transport mode).
- [ ] Commit "Docs and help for local links".

### Task 17: The clipboard probe and both click-throughs in local mode

- [ ] `scripts/ux/clipboard-probe.mjs`: the 23.09 spike made permanent (uses `scripts/ux/browser.mjs`); it builds
  a real bundle with `src/link/bundle.ts` through Vite (the way `scripts/ux/panes.mjs` imports `/src/...`),
  writes it with `beginClipboardWrite` and `copyBundleNow` in one page and reads it with `bundleFromPaste` in a
  second page of the same origin; fails unless both paths round-trip a 5 MB bundle and the plain text is the
  sentence. `npm run ux:clipboard` in `package.json`; not part of `check` (it needs the Playwright browser).
- [ ] Both click-through suites run every state once in each transport (a `describe.each(["relay", "local"])`
  around the existing states, the local run seeding no key).
- [ ] Gates: `npm run check`, `npm run ux:sweep`, `npm run ux:check`, `npm run ux:clipboard`, all green.
- [ ] Commit "Clipboard probe; click-through in both transports".

### Task 18: Real-Office checklist rows

- [ ] Add a section to `tasks/launch-check.md`: Mac Excel -> PowerPoint, local mode: export one range, paste,
  insert; change it, Copy all, paste, the deck updates; paste the bundle onto a slide gives the sentence; the
  Copy for PowerPoint bar appears only if the first write was refused (note which); reopen PowerPoint, the Inbox
  still lists what waited; Clear pasted links. Windows: the same rows, later.
- [ ] Commit "Launch check: local links rows".

---

## Self-review

- Spec coverage: carrier (Tasks 1, 8), setting and default (9, 12, 13), Excel writer and revs (2, 3, 10, 11),
  PowerPoint store, persistence, statuses, paste, revert, Sent column, Clear (4, 5, 6, 7, 13, 14, 15),
  security and privacy (16), errors table (1, 11, 14, 15), testing (every task, 17, 18), rollout (slice order).
- Types used across tasks: `Bundle`, `BundleLink`, `BundleInboxRow`, `BundleRead` (T1); `LOCAL_REV_BASE`,
  `nextLocalRev`, `previousRevOf`, `localWorkspace` (T2); `LocalCollector` (T3); `LinkStatus.notPasted`,
  `RelayStatus.local`, `OmittedReason.notPasted` (T4); `LocalPersistence`, `StoredLocalLink`,
  `StoredLocalInbox`, `LocalStore`, `IngestResult`, `NOT_PASTED_HERE` (T5, T6); `beginClipboardWrite`,
  `copyBundleNow`, `bundleFromPaste` (T8); `LinkTransport`, `loadTransport`, `saveTransport` (T9);
  `PushOptions` (T10); `LinksTransport` (T11); `PptTransport` (T13); `PasteSummary`, `pasteLinks`,
  `summarizePaste` (T14). Names match between producer and consumer tasks.
