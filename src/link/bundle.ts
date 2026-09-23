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
