// Link core: the JSON shapes carried by the Excel tag, the workbook registry,
// the relay payload and the PowerPoint inbox item, plus their codecs. Every
// decoder validates its shape before trusting it - garbage in never becomes a
// typed value out.

export type LinkId = string; // 32 lowercase hex chars
export type LinkKind = "range" | "chart";

export interface Source {
  workbook: string;
  sheet: string;
  ref: string;
  anchor: string;
}

export interface LinkTag {
  v: 1;
  id: LinkId;
  kind: LinkKind;
  rev: number;
  src: Source;
  pushedAt: string;
}

export interface RegistryEntry {
  id: LinkId;
  kind: LinkKind;
  anchor: string;
  label: string;
  token: string;
  createdAt: string;
  lastPushedAt: string | null;
  rev: number;
}

export interface Registry {
  v: 1;
  links: RegistryEntry[];
}

export interface Payload {
  v: 1;
  kind: "picture";
  mime: "image/png";
  width: number;
  height: number;
  png: string;
  src: Source;
  pushedAt: string;
  hash: string;
}

export interface InboxItem {
  id: LinkId;
  token: string;
  kind: LinkKind;
  label: string;
  src: Source;
  createdAt: string;
}

export const TAG_LINK = "SMT_LINK";
export const TAG_KEY = "SMT_KEY";
export const REGISTRY_SETTING = "SMT_LINKS";
export const ANCHOR_PREFIX = "SMT_LINK_";
export const TAG_VALUE_MAX = 2048;

const LINK_ID_PATTERN = /^[0-9a-f]{32}$/;
const LINK_ID_BYTES = 16;
const ANCHOR_ID_CHARS = 8;
const HEX_RADIX = 16;

export function newLinkId(random: (n: number) => Uint8Array): LinkId {
  const bytes = random(LINK_ID_BYTES);
  return Array.from(bytes, (byte) =>
    byte.toString(HEX_RADIX).padStart(2, "0"),
  ).join("");
}

export function isLinkId(value: string): boolean {
  return LINK_ID_PATTERN.test(value);
}

export function anchorName(id: LinkId): string {
  return ANCHOR_PREFIX + id.slice(0, ANCHOR_ID_CHARS);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isSource(value: unknown): value is Source {
  return (
    isRecord(value) &&
    typeof value.workbook === "string" &&
    typeof value.sheet === "string" &&
    typeof value.ref === "string" &&
    typeof value.anchor === "string"
  );
}

function isKind(value: unknown): value is LinkKind {
  return value === "range" || value === "chart";
}

function isTag(value: unknown): value is LinkTag {
  return (
    isRecord(value) &&
    value.v === 1 &&
    typeof value.id === "string" &&
    isKind(value.kind) &&
    typeof value.rev === "number" &&
    isSource(value.src) &&
    typeof value.pushedAt === "string"
  );
}

function isEntry(value: unknown): value is RegistryEntry {
  return (
    isRecord(value) &&
    typeof value.id === "string" &&
    isKind(value.kind) &&
    typeof value.anchor === "string" &&
    typeof value.label === "string" &&
    typeof value.token === "string" &&
    typeof value.createdAt === "string" &&
    (value.lastPushedAt === null || typeof value.lastPushedAt === "string") &&
    typeof value.rev === "number"
  );
}

function isRegistry(value: unknown): value is Registry {
  return (
    isRecord(value) &&
    value.v === 1 &&
    Array.isArray(value.links) &&
    value.links.every(isEntry)
  );
}

function isPayload(value: unknown): value is Payload {
  return (
    isRecord(value) &&
    value.v === 1 &&
    value.kind === "picture" &&
    value.mime === "image/png" &&
    typeof value.width === "number" &&
    typeof value.height === "number" &&
    typeof value.png === "string" &&
    isSource(value.src) &&
    typeof value.pushedAt === "string" &&
    typeof value.hash === "string"
  );
}

function isInboxItem(value: unknown): value is InboxItem {
  return (
    isRecord(value) &&
    typeof value.id === "string" &&
    typeof value.token === "string" &&
    isKind(value.kind) &&
    typeof value.label === "string" &&
    isSource(value.src) &&
    typeof value.createdAt === "string"
  );
}

export function encodeTag(tag: LinkTag): string {
  const json = JSON.stringify(tag);
  if (json.length > TAG_VALUE_MAX) {
    throw new Error(`encodeTag: tag exceeds ${TAG_VALUE_MAX} chars`);
  }
  return json;
}

export function decodeTag(value: string | null | undefined): LinkTag | null {
  if (value === null || value === undefined) return null;
  try {
    const parsed: unknown = JSON.parse(value);
    return isTag(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

export function encodeRegistry(registry: Registry): string {
  return JSON.stringify(registry);
}

// Null says "this is not a registry we understand" - unparseable, or a shape a
// future schema wrote. A caller that is about to write the setting back must
// tell that apart from an absent one, or it erases every link record.
export function tryDecodeRegistry(
  value: string | null | undefined,
): Registry | null {
  if (value === null || value === undefined) return null;
  try {
    const parsed: unknown = JSON.parse(value);
    return isRegistry(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

export function emptyRegistry(): Registry {
  return { v: 1, links: [] };
}

// Forgiving by design, for readers that only need a list to show.
export function decodeRegistry(value: string | null | undefined): Registry {
  return tryDecodeRegistry(value) ?? emptyRegistry();
}

function decodeJson(bytes: Uint8Array, stage: string): unknown {
  try {
    return JSON.parse(new TextDecoder().decode(bytes));
  } catch {
    throw new Error(`${stage}: not valid JSON`);
  }
}

export function encodePayload(payload: Payload): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(payload));
}

export function decodePayload(bytes: Uint8Array): Payload {
  const parsed = decodeJson(bytes, "decodePayload");
  if (!isPayload(parsed)) throw new Error("decodePayload: not a link payload");
  return parsed;
}

export function encodeInboxItem(item: InboxItem): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(item));
}

export function decodeInboxItem(bytes: Uint8Array): InboxItem {
  const parsed = decodeJson(bytes, "decodeInboxItem");
  if (!isInboxItem(parsed)) {
    throw new Error("decodeInboxItem: not an inbox item");
  }
  return parsed;
}

export function sourceLabel(src: Source, kind: LinkKind): string {
  return kind === "range"
    ? `${src.sheet}!${src.ref}`
    : `${src.sheet}: ${src.ref}`;
}
