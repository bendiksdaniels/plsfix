// Attacks the relay-failure wording shared by pushOne and publish's catch
// (relayFailureReason, src/excel/link-record.ts) across push, auto-push and
// export: a 429, 507 or other 5xx must never reach the modeller as the
// relay's raw status line, e.g. "relay PUT /api/links/<id>: 507 storage full".

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../../src/excel/links";
import type * as WatchModule from "../../src/excel/link-watch";
import { RelayError } from "../../src/link/relay";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../../src/link/workspace";
import { virtualClock, type VirtualClock } from "../clock";
import { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "../fakehost";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let watch: typeof WatchModule;
let helpers: FakeHelpers;
let relay: FakeRelay;
let ws: Workspace;

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

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  host.workbook.fileUrl = "/Users/daniel/Models/Model_v4.xlsx";
  relay = new FakeRelay();
  ws = await createWorkspace(memoryStore());
  links = await import("../../src/excel/links");
  watch = await import("../../src/excel/link-watch");
  helpers.seed("Model!B4", [[1, 2]]);
  helpers.select("Model!B4:C4");
});
afterEach(() => uninstallFakeHost());

const CASES: [string, RelayError, string][] = [
  [
    "429 (the write rate limit)",
    new RelayError("server", "relay PUT /api/links/x: 429", 429),
    "The relay is busy. Try again in a minute.",
  ],
  [
    "507 (the relay is full)",
    new RelayError("server", "relay PUT /api/links/x: 507 storage full", 507),
    "The relay is full. Ask for space to be cleared.",
  ],
  [
    "502 (any other fault)",
    new RelayError("server", "relay PUT /api/links/x: 502 bad gateway", 502),
    "The relay had a problem. Try again in a minute.",
  ],
];

describe("Push / Push all: a relay refusal reaches the failure list in words", () => {
  it.each(CASES)(
    "says %s in the pane's own words",
    async (_name, error, sentence) => {
      const { id } = await links.exportSelection(ws, relay);
      relay.putLink = () => Promise.reject(error);

      const summary = await links.pushLinks("all", relay);

      expect(summary).toEqual({
        pushed: 0,
        missing: 0,
        failed: 1,
        failures: [`Model!B4:C4: ${sentence}`],
      });
      expect(relay.links.get(id)!.rev).toBe(1);
    },
  );

  it("still keeps the 413 sentence worded for the Excel side, not PowerPoint's", async () => {
    await links.exportSelection(ws, relay);
    relay.putLink = () =>
      Promise.reject(new RelayError("tooLarge", "413 payload too large", 413));

    const summary = await links.pushLinks("all", relay);

    expect(summary.failures).toEqual([
      "Model!B4:C4: That export is too big to send. Export a smaller range.",
    ]);
  });

  it("leaves an error the relay never sent to travel as it said it", async () => {
    await links.exportSelection(ws, relay);
    relay.putLink = () => Promise.reject(new Error("ItemNotFound"));

    const summary = await links.pushLinks("all", relay);

    expect(summary.failures).toEqual(["Model!B4:C4: ItemNotFound"]);
  });
});

describe("Export: a relay refusal on the very first push reaches the caller in words", () => {
  // Bug 2 was only half closed: pushOne (above) worded these four, but
  // publish's own catch (src/excel/link-record.ts, what exportSelection
  // throws through when the export's OWN first push fails) still threw the
  // raw relay error, e.g. "export Model!B4:C4: relay PUT /api/links/x: 507".
  it.each(CASES)(
    "says %s in the pane's own words, not the raw status line",
    async (_name, error, sentence) => {
      relay.putLink = () => Promise.reject(error);

      await expect(links.exportSelection(ws, relay)).rejects.toThrow(
        `export Model!B4:C4: ${sentence}`,
      );
    },
  );

  it("still keeps the 413 sentence worded for the Excel side, not PowerPoint's", async () => {
    relay.putLink = () =>
      Promise.reject(new RelayError("tooLarge", "413 payload too large", 413));

    await expect(links.exportSelection(ws, relay)).rejects.toThrow(
      "export Model!B4:C4: That export is too big to send. Export a smaller range.",
    );
  });

  it("leaves an error the relay never sent to travel as it said it", async () => {
    relay.putLink = () => Promise.reject(new Error("ItemNotFound"));

    await expect(links.exportSelection(ws, relay)).rejects.toThrow(
      "export Model!B4:C4: ItemNotFound",
    );
  });
});

describe("auto-push: the same refusal, reported through notify", () => {
  let clock: VirtualClock;
  let notes: string[];

  beforeEach(() => {
    clock = virtualClock();
    notes = [];
  });

  async function armed(): Promise<void> {
    await links.exportSelection(ws, relay);
    await watch.setAutoPush(true, relay, (message) => notes.push(message), {
      clock,
    });
  }

  it.each(CASES)(
    "turns %s into the same plain sentence auto-push shows",
    async (_name, error, sentence) => {
      await armed();
      relay.putLink = () => Promise.reject(error);

      await helpers.fireChanged("Model", "C4");
      clock.advance(watch.AUTOPUSH_DELAY_MS);

      await vi.waitFor(() => {
        expect(notes).toEqual([`Auto-push failed: Model!B4:C4: ${sentence}`]);
      });
    },
  );
});
