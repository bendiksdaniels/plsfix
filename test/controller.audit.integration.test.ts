// Controller follow-ups after the audit reviews: an export the relay refuses
// as too large (413) reaches the modeller as a sentence with a way out, on the
// Excel side where the push actually happens.
import { beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import { createWorkspace, type KeyStore } from "../src/link/workspace";
import { RelayError } from "../src/link/relay";
import { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "./fakehost";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let helpers: FakeHelpers;
let relay: FakeRelay;

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
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  host.workbook.fileUrl = "/Users/daniel/Models/Model_v4.xlsx";
  relay = new FakeRelay();
  links = await import("../src/excel/links");
});

describe("a push the relay refuses as too large", () => {
  it("says so in words with a way out", async () => {
    const ws = await createWorkspace(memoryStore());
    helpers.seed("Model!B4", [[1, 2]]);
    helpers.select("Model!B4:C4");
    await links.exportSelection(ws, relay);

    relay.putLink = () =>
      Promise.reject(
        new RelayError("tooLarge", "relay PUT /api/links/x: 413", 413),
      );

    const summary = await links.pushLinks("all", relay);

    expect(summary.failed).toBe(1);
    expect(summary.failures[0]).toContain(
      "That export is too big to send. Export a smaller range.",
    );
  });
});
