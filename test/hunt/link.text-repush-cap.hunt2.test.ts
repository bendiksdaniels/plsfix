// Pass-2 finding: TEXT_MAX_CHARS was checked only at export time; a text
// link's cell can grow past 500 chars afterwards, and payloadOf never
// re-checked it on a later push, unlike the table cap. Pin: every over-cap
// payload is refused before sealing, on every render, not just the first.

import fc from "fast-check";
import { afterEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../../src/excel/links";
import { TEXT_MAX_CHARS, TEXT_TOO_LONG } from "../../src/link/model";
import { createWorkspace, type KeyStore } from "../../src/link/workspace";
import { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "../fakehost";

enableStrictLoadSemantics();

const SEED = 20260927;
const RUNS = 40;

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

async function linkedShortText(helpers: FakeHelpers): Promise<{
  links: typeof LinksModule;
  relay: FakeRelay;
  id: string;
  registryBefore: string | null;
}> {
  const links = await import("../../src/excel/links");
  helpers.seed("Model!B4", [["short"]]);
  helpers.select("Model!B4");
  const relay = new FakeRelay();
  const ws = await createWorkspace(memoryStore());
  const { id } = await links.exportSelectionAsText(ws, relay);
  return { links, relay, id, registryBefore: helpers.setting("PLSFIX_LINKS") };
}

describe("text cap: a source that grew past 500 characters since it was linked", () => {
  afterEach(() => uninstallFakeHost());

  it("refuses a later push with the same sentence export uses, and never calls the relay", async () => {
    const { helpers } = installFakeHost({ sheets: ["Model"] });
    const { links, relay, registryBefore } = await linkedShortText(helpers);

    helpers.seed("Model!B4", [["x".repeat(TEXT_MAX_CHARS + 1)]]);
    const putLinkSpy = vi.spyOn(relay, "putLink");

    const summary = await links.pushLinks("all", relay);

    expect(summary).toMatchObject({ pushed: 0, missing: 0, failed: 1 });
    expect(summary.failures[0]).toContain(TEXT_TOO_LONG);
    expect(putLinkSpy).not.toHaveBeenCalled();
    // Nothing half-written: the registry is byte-for-byte what export left,
    // not a revision bumped for a push that never landed.
    expect(helpers.setting("PLSFIX_LINKS")).toBe(registryBefore);
  });

  it("refuses on Push all mixed with a healthy link, and only fails the one over cap", async () => {
    const { helpers } = installFakeHost({ sheets: ["Model"] });
    const { links, relay } = await linkedShortText(helpers);
    helpers.seed("Model!C4", [["also short"]]);
    helpers.select("Model!C4");
    await links.exportSelectionAsText(
      await createWorkspace(memoryStore()),
      relay,
    );

    helpers.seed("Model!B4", [["y".repeat(TEXT_MAX_CHARS + 50)]]);
    const summary = await links.pushLinks("all", relay);

    expect(summary).toMatchObject({ pushed: 1, missing: 0, failed: 1 });
    expect(summary.failures[0]).toContain(TEXT_TOO_LONG);
  });

  it("refuses over any generated over-cap length, unicode included, before the relay is ever called", async () => {
    const overCapLengthArb = fc.integer({
      min: TEXT_MAX_CHARS + 1,
      max: TEXT_MAX_CHARS + 2000,
    });
    await fc.assert(
      fc.asyncProperty(overCapLengthArb, async (length) => {
        uninstallFakeHost();
        const { helpers } = installFakeHost({ sheets: ["Model"] });
        try {
          const { links, relay, registryBefore } =
            await linkedShortText(helpers);
          helpers.seed("Model!B4", [["é".repeat(length)]]);
          const putLinkSpy = vi.spyOn(relay, "putLink");

          const summary = await links.pushLinks("all", relay);

          expect(summary.pushed).toBe(0);
          expect(summary.failed).toBe(1);
          expect(summary.failures[0]).toContain(TEXT_TOO_LONG);
          expect(putLinkSpy).not.toHaveBeenCalled();
          expect(helpers.setting("PLSFIX_LINKS")).toBe(registryBefore);
        } finally {
          uninstallFakeHost();
        }
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
