// Pass-2 findings: two relay refusals reached a modeller as a raw status
// line. (1) relayReason had no 408 case, though the server's own timeout
// answers it for real. (2) forget() (Remove) never called relayReason at
// all, unlike every other relay-touching flow.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { forget } from "../../src/excel/link-anchors";
import type * as LinksModule from "../../src/excel/links";
import { newToken } from "../../src/link/crypto";
import type { RegistryEntry } from "../../src/link/model";
import { RelayError, type RelayApi } from "../../src/link/relay";
import { relayReason } from "../../src/link/relay-reason";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../../src/link/workspace";
import { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "../fakehost";

enableStrictLoadSemantics();

function memoryStore(): KeyStore {
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

// A raw relay line a modeller cannot act on: the shape every real
// RelayClient failure and every fake in this suite constructs.
function rawLine(status: number): string {
  return `relay PUT /api/links/x: ${String(status)}`;
}

describe("relayReason: 408 (the server's own request timeout)", () => {
  it("words a 408 the same family as 429/507/5xx, not the raw status line", () => {
    const error = new RelayError("server", rawLine(408), 408);
    expect(relayReason(error)).toBe(
      "The link relay did not answer in time. Try again.",
    );
  });
});

describe("forget(): a 429/507/5xx DELETE failure reaches the pane in words", () => {
  const entry: RegistryEntry = {
    id: "a".repeat(32),
    kind: "range",
    anchor: "PLSFIX_LINK_a",
    label: "Model!A1",
    token: newToken(),
    createdAt: "2026-01-01T00:00:00.000Z",
    lastPushedAt: null,
    rev: 1,
  };

  it.each([
    [
      "429",
      new RelayError("server", rawLine(429), 429),
      "The relay is busy. Try again in a minute.",
    ],
    [
      "507",
      new RelayError("server", rawLine(507), 507),
      "The relay is full. Ask for space to be cleared.",
    ],
    [
      "503",
      new RelayError("server", rawLine(503), 503),
      "The relay had a problem. Try again in a minute.",
    ],
    [
      "408",
      new RelayError("server", rawLine(408), 408),
      "The link relay did not answer in time. Try again.",
    ],
  ] as [string, RelayError, string][])(
    "says %s in the pane's own words, not the raw status line",
    async (_name, error, sentence) => {
      const relay = {
        deleteLink: () => Promise.reject(error),
      } as unknown as RelayApi;

      await expect(forget(entry, relay)).rejects.toThrow(
        `remove Model!A1: ${sentence}; nothing was removed, try again`,
      );
    },
  );

  it("still reports an error the relay never sent as it said it (unchanged)", async () => {
    const relay = {
      deleteLink: () => Promise.reject(new Error("ItemNotFound")),
    } as unknown as RelayApi;

    await expect(forget(entry, relay)).rejects.toThrow(
      "remove Model!A1: ItemNotFound; nothing was removed, try again",
    );
  });
});

describe("end to end: removeLink surfaces the same plain sentence", () => {
  let links: typeof LinksModule;
  let helpers: FakeHelpers;
  let relay: FakeRelay;
  let ws: Workspace;

  beforeEach(async () => {
    vi.resetModules();
    uninstallFakeHost();
    const host = installFakeHost({ sheets: ["Model"] });
    helpers = host.helpers;
    relay = new FakeRelay();
    ws = await createWorkspace(memoryStore());
    links = await import("../../src/excel/links");
    helpers.seed("Model!B4", [[1]]);
    helpers.select("Model!B4");
  });
  afterEach(() => uninstallFakeHost());

  it("removeLink: a 429 reaches the caller worded, and the entry stays until it can be removed", async () => {
    const { id } = await links.exportSelection(ws, relay);
    const realDeleteLink = relay.deleteLink.bind(relay);
    relay.deleteLink = () =>
      Promise.reject(new RelayError("server", rawLine(429), 429));

    await expect(links.removeLink(id, relay)).rejects.toThrow(
      "remove Model!B4: The relay is busy. Try again in a minute.; nothing was removed, try again",
    );

    const rows = await links.listWorkbookLinks();
    expect(rows).toHaveLength(1);
    expect(rows[0]!.entry.id).toBe(id);

    // The next attempt, against a relay that now answers, works cleanly.
    relay.deleteLink = realDeleteLink;
    await links.removeLink(id, relay);
    expect(await links.listWorkbookLinks()).toHaveLength(0);
  });
});
