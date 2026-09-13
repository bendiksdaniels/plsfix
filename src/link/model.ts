// Link core: the JSON shapes carried by the Excel tag, the workbook registry,
// the relay payload (a picture, a table or a text) and the inbox item, plus
// their codecs. Every decoder validates its shape before trusting it - garbage
// in never becomes a typed value out.

import { guardChart } from "./chart-guard";
import type { ChartData } from "./chart-model";
import {
  cleanProjectList,
  optionalProject,
  readProject,
  withProject,
} from "./project";

export type LinkId = string; // 32 lowercase hex chars
export type LinkKind = "range" | "chart" | "table" | "text";

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
  project?: string;
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
  project?: string;
}

export interface Registry {
  v: 1;
  links: RegistryEntry[];
  projects?: string[];
  activeProject?: string;
}

export interface PicturePayload {
  v: 1;
  kind: "picture";
  mime: "image/png";
  width: number;
  height: number;
  png: string;
  src: Source;
  pushedAt: string;
  hash: string;
  // A chart link also ships what the chart shows, so a slide that can draw
  // native shapes does; the picture stays the fallback for every other host.
  chart?: ChartData;
  // Or, when Excel could not describe the chart, why not: the sentence the
  // slide shows beside the picture it inserts instead.
  chartIssue?: string;
}

// One cell of a table, as small as it can be said: `t` is the text Excel
// displays, and every other key is omitted when the cell wears the default, so
// a plain grid travels as little more than its own words. `f` absent means no
// fill at all, not white.
export interface TableCell {
  t: string;
  b?: true;
  i?: true;
  c?: string;
  f?: string;
  a?: "l" | "c" | "r";
  z?: number;
}

// The grid itself: `rows` and `cols` say what `cells` must be, and `widths`
// carries one column width in points per column.
export interface TablePayload {
  v: 1;
  kind: "table";
  rows: number;
  cols: number;
  cells: TableCell[][];
  widths: number[];
  src: Source;
  pushedAt: string;
  hash: string;
}

// One cell's displayed text, nothing else: the slide decides the font. The
// hash is over the text, so an unchanged cell is an unchanged link.
export interface TextPayload {
  v: 1;
  kind: "text";
  text: string;
  src: Source;
  pushedAt: string;
  hash: string;
}

export type Payload = PicturePayload | TablePayload | TextPayload;

// A cell longer than this is a paragraph: a picture or a table says it better.
export const TEXT_MAX_CHARS = 500;
export const TEXT_TOO_LONG =
  `Text links carry up to ${String(TEXT_MAX_CHARS)} characters. ` +
  `Export a longer cell as a picture.`;

// A native PowerPoint table is written cell by cell, so its size is what an
// insert costs: past this, a picture is the honest answer.
export const TABLE_MAX_ROWS = 60;
export const TABLE_MAX_COLS = 20;
export const TABLE_TOO_BIG =
  `Tables go up to ${String(TABLE_MAX_ROWS)} rows and ` +
  `${String(TABLE_MAX_COLS)} columns; export a picture for more.`;

export function overTableCap(rows: number, cols: number): boolean {
  return rows > TABLE_MAX_ROWS || cols > TABLE_MAX_COLS;
}

// What one repaint round trip carries: the base64 of a picture, or the JSON of
// a table's cells - the two are the payload, everything else is a header.
export function payloadBytes(payload: Payload): number {
  if (payload.kind === "text")
    return new TextEncoder().encode(payload.text).length;
  return payload.kind === "picture"
    ? payload.png.length
    : JSON.stringify(payload.cells).length;
}

export interface InboxItem {
  id: LinkId;
  token: string;
  kind: LinkKind;
  label: string;
  src: Source;
  createdAt: string;
  project?: string;
}

export const TAG_LINK = "PLSFIX_LINK";
export const TAG_KEY = "PLSFIX_KEY";
export const REGISTRY_SETTING = "PLSFIX_LINKS";
export const ANCHOR_PREFIX = "PLSFIX_LINK_";
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

const KINDS: LinkKind[] = ["range", "chart", "table", "text"];

function isKind(value: unknown): value is LinkKind {
  return KINDS.some((kind) => kind === value);
}

function isTag(value: unknown): value is LinkTag {
  return (
    isRecord(value) &&
    value.v === 1 &&
    typeof value.id === "string" &&
    isKind(value.kind) &&
    typeof value.rev === "number" &&
    isSource(value.src) &&
    typeof value.pushedAt === "string" &&
    optionalProject(value.project)
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
    typeof value.rev === "number" &&
    optionalProject(value.project)
  );
}

function isRegistry(value: unknown): value is Registry {
  return (
    isRecord(value) &&
    value.v === 1 &&
    Array.isArray(value.links) &&
    value.links.every(isEntry) &&
    isOptional(
      value.projects,
      (entry) =>
        Array.isArray(entry) && entry.every((name) => typeof name === "string"),
    ) &&
    optionalProject(value.activeProject)
  );
}

function isPicturePayload(value: unknown): value is PicturePayload {
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
    typeof value.hash === "string" &&
    isOptional(value.chartIssue, (entry) => typeof entry === "string")
  );
}

function isOptional(value: unknown, ok: (entry: unknown) => boolean): boolean {
  return value === undefined || ok(value);
}

// A key that is present must carry a value the writer would have written: an
// absent bold and a `false` bold are not the same thing here, because the
// encoder omits every default rather than spelling it out.
function isTableCell(value: unknown): value is TableCell {
  return (
    isRecord(value) &&
    typeof value.t === "string" &&
    isOptional(value.b, (entry) => entry === true) &&
    isOptional(value.i, (entry) => entry === true) &&
    isOptional(value.c, (entry) => typeof entry === "string") &&
    isOptional(value.f, (entry) => typeof entry === "string") &&
    isOptional(value.z, (entry) => typeof entry === "number") &&
    isOptional(
      value.a,
      (entry) => entry === "l" || entry === "c" || entry === "r",
    )
  );
}

// The counts are not a second opinion about the grid: a payload whose cells do
// not make exactly rows x cols is refused rather than half-drawn.
function isCellGrid(value: unknown, rows: number, cols: number): boolean {
  return (
    Array.isArray(value) &&
    value.length === rows &&
    value.every(
      (row) =>
        Array.isArray(row) && row.length === cols && row.every(isTableCell),
    )
  );
}

function isTablePayload(value: unknown): value is TablePayload {
  return (
    isRecord(value) &&
    value.v === 1 &&
    value.kind === "table" &&
    typeof value.rows === "number" &&
    typeof value.cols === "number" &&
    isCellGrid(value.cells, value.rows, value.cols) &&
    Array.isArray(value.widths) &&
    value.widths.length === value.cols &&
    value.widths.every((width) => typeof width === "number") &&
    isSource(value.src) &&
    typeof value.pushedAt === "string" &&
    typeof value.hash === "string"
  );
}

function isTextPayload(value: unknown): value is TextPayload {
  return (
    isRecord(value) &&
    value.v === 1 &&
    value.kind === "text" &&
    typeof value.text === "string" &&
    isSource(value.src) &&
    typeof value.pushedAt === "string" &&
    typeof value.hash === "string"
  );
}

function isPayload(value: unknown): value is Payload {
  return (
    isPicturePayload(value) || isTablePayload(value) || isTextPayload(value)
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
    typeof value.createdAt === "string" &&
    optionalProject(value.project)
  );
}

export function encodeTag(tag: LinkTag): string {
  const json = JSON.stringify(withProject(tag));
  if (json.length > TAG_VALUE_MAX) {
    throw new Error(`encodeTag: tag exceeds ${TAG_VALUE_MAX} chars`);
  }
  return json;
}

export function decodeTag(value: string | null | undefined): LinkTag | null {
  if (value === null || value === undefined) return null;
  try {
    const parsed: unknown = JSON.parse(value);
    return isTag(parsed) ? withProject(parsed) : null;
  } catch {
    return null;
  }
}

export function encodeRegistry(registry: Registry): string {
  return JSON.stringify(cleanRegistry(registry));
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
    return isRegistry(parsed) ? cleanRegistry(parsed) : null;
  } catch {
    return null;
  }
}

function cleanRegistry(registry: Registry): Registry {
  const links = registry.links.map(withProject);
  const projects = cleanProjectList(registry.projects);
  const activeProject = readProject(registry.activeProject);
  return {
    v: 1,
    links,
    ...(projects.length > 0 ? { projects } : {}),
    ...(activeProject ? { activeProject } : {}),
  };
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
  const parsed = guardChart(decodeJson(bytes, "decodePayload"));
  if (!isPayload(parsed)) throw new Error("decodePayload: not a link payload");
  return parsed;
}

export function encodeInboxItem(item: InboxItem): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(withProject(item)));
}

export function decodeInboxItem(bytes: Uint8Array): InboxItem {
  const parsed = decodeJson(bytes, "decodeInboxItem");
  if (!isInboxItem(parsed)) {
    throw new Error("decodeInboxItem: not an inbox item");
  }
  return withProject(parsed);
}

// A table names its range and says so: two links can point at the same cells,
// one as a picture and one as a table, and the list has to tell them apart.
export function sourceLabel(src: Source, kind: LinkKind): string {
  if (kind === "chart") return `${src.sheet}: ${src.ref}`;
  const range = `${src.sheet}!${src.ref}`;
  if (kind === "table") return `${range} table`;
  return kind === "text" ? `${range} text` : range;
}
