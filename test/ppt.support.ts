// Shared PowerPoint fake-host setup for the ppt.*.integration.test.ts suites:
// the workspace key store, the relay-seeding helpers (seedLink and seedTable,
// pushAgain and pushTable), the source-cell constant they publish under, and
// the per-test boot (bootPpt) that installs a fresh fake deck and relay.

import { vi } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import {
  encodePayload,
  newLinkId,
  type InboxItem,
  type Payload,
  type TableCell,
  type TablePayload,
} from "../src/link/model";
import type { KeyStore } from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import {
  installFakePpt,
  uninstallFakePpt,
  type FakePresentation,
  type FakePptHelpers,
} from "./fakeppt";
import type * as LinksModule from "../src/ppt/links";

export const src = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

let relay: FakeRelay;

export function memoryStore(): KeyStore {
  const map = new Map<string, string>();
  return {
    get: async (k) => map.get(k) ?? null,
    set: async (k, v) => {
      map.set(k, v);
    },
    remove: async (k) => {
      map.delete(k);
    },
  };
}

function newId(): string {
  return newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
}

async function publish(
  id: string,
  token: string,
  payload: Payload,
): Promise<void> {
  const keys = await deriveLinkKeys(token);
  await relay.putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
  );
}

function tablePayload(
  cells: TableCell[][],
  widths: number[],
  hash: string,
): TablePayload {
  return {
    v: 1,
    kind: "table",
    rows: cells.length,
    cols: widths.length,
    cells,
    widths,
    src,
    pushedAt: new Date().toISOString(),
    hash,
  };
}

// A table export waiting in the inbox, sealed the way Excel would have sealed
// it: the label is the one sourceLabel gives a table link.
export async function seedTable(
  cells: TableCell[][],
  widths: number[],
): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  const payload = tablePayload(cells, widths, "0".repeat(64));
  await publish(id, token, payload);
  return {
    id,
    token,
    kind: "table",
    label: `${src.sheet}!${src.ref} table`,
    src: payload.src,
    createdAt: payload.pushedAt,
  };
}

export async function pushTable(
  item: InboxItem,
  cells: TableCell[][],
  widths: number[],
): Promise<void> {
  await publish(
    item.id,
    item.token,
    tablePayload(cells, widths, "1".repeat(64)),
  );
}

export async function seedLink(
  png: string,
  workbook = src.workbook,
): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  const payload: Payload = {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png,
    src: { ...src, workbook },
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };
  await publish(id, token, payload);
  return {
    id,
    token,
    kind: "range",
    label: "Model!B4:F12",
    src: payload.src,
    createdAt: payload.pushedAt,
  };
}

export async function pushAgain(
  item: InboxItem,
  png: string,
  workbook = src.workbook,
): Promise<void> {
  const payload: Payload = {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 1,
    height: 1,
    png,
    src: { ...src, workbook },
    pushedAt: new Date().toISOString(),
    hash: "1".repeat(64),
  };
  await publish(item.id, item.token, payload);
}

// The per-test boot: a fresh fake deck (3 slides, slide 0 selected) and a
// fresh relay, the same sequence every ppt.*.integration.test.ts beforeEach
// used to run inline.
export async function bootPpt(): Promise<{
  links: typeof LinksModule;
  presentation: FakePresentation;
  helpers: FakePptHelpers;
  relay: FakeRelay;
}> {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 3 });
  const presentation = host.presentation;
  const helpers = host.helpers;
  relay = new FakeRelay();
  const links = await import("../src/ppt/links");
  helpers.selectSlide(presentation.slides[0]!.id);
  return { links, presentation, helpers, relay };
}
