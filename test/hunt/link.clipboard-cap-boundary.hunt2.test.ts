// Pass-2: src/pane/links-transport.ts's copy() enforces BUNDLE_MAX_CHARS
// exactly - a bundle at the cap copies, one character more is refused.
// Split out of link.bundle-payload.properties for the line cap.
// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { BUNDLE_MAX_CHARS, encodeBundle } from "../../src/link/bundle";
import { LocalCollector } from "../../src/link/local-collector";
import { installLinksTransport } from "../../src/pane/links-transport";
import type { KeyStore } from "../../src/link/workspace";
import type { Toast } from "../../src/ui/toast";

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

describe("clipboard bundle cap: enforced exactly at the character boundary", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  function paneRoot(): Document {
    document.body.innerHTML = readFileSync(
      join(process.cwd(), "taskpane.html"),
      "utf8",
    );
    return document;
  }

  function fakeToast(): Toast {
    return { show: () => undefined };
  }

  // How many bytes a single link's blob needs so that copy()'s own
  // encodeBundle(collector.bundle()) lands at exactly `target` characters:
  // measured against a real LocalCollector (never guessed), because rev's
  // digit width and sentAt's real Date.now() width both count. base64url
  // never emits length = 4k+1 for any byte count, so a target one char
  // short of a whole 4-byte group is made reachable by asking putLink for a
  // rev one digit wider instead (LOCAL_REV_BASE has 13 digits; a currentRev
  // of 10^13 forces a 14-digit one) - both measured, neither guessed.
  const WIDER_REV = 10_000_000_000_000; // 14 digits once nextLocalRev adds 1
  async function measuredLength(
    id: string,
    blobBytes: number,
    widerRev: boolean,
  ): Promise<number> {
    const probe = new LocalCollector();
    await probe.putLink(
      id,
      "auth",
      new Uint8Array(blobBytes),
      widerRev ? WIDER_REV : 0,
    );
    return encodeBundle(probe.bundle()).length;
  }

  async function blobBytesForLength(
    id: string,
    target: number,
  ): Promise<{ blobBytes: number; widerRev: boolean }> {
    let widerRev = false;
    let base = await measuredLength(id, 0, widerRev);
    let remainder = target - base;
    if (remainder % 4 === 1) {
      widerRev = true;
      base = await measuredLength(id, 0, widerRev);
      remainder = target - base;
    }
    const blobBytes =
      remainder % 4 === 0
        ? (remainder / 4) * 3
        : remainder % 4 === 2
          ? Math.floor(remainder / 4) * 3 + 1
          : Math.floor(remainder / 4) * 3 + 2;
    return { blobBytes, widerRev };
  }

  it("accepts a bundle of exactly BUNDLE_MAX_CHARS and refuses one character more", async () => {
    vi.stubGlobal(
      "ClipboardItem",
      class {
        constructor(readonly items: Record<string, Promise<Blob>>) {}
      },
    );
    vi.stubGlobal("navigator", {
      // Drains both clipboard flavors either way (resolved for the at-cap
      // copy, rejected for the over-cap one, once pending.reject fires):
      // allSettled never itself rejects, so neither promise is left
      // unhandled.
      clipboard: {
        write: async (items: { items: Record<string, Promise<Blob>> }[]) => {
          await Promise.allSettled(Object.values(items[0]!.items));
        },
      },
    });

    const id = "a".repeat(32);
    const atCap = await blobBytesForLength(id, BUNDLE_MAX_CHARS);
    const overCap = await blobBytesForLength(id, BUNDLE_MAX_CHARS + 1);

    const t = await installLinksTransport({
      root: paneRoot(),
      keyStore: memoryStore(),
      toast: fakeToast(),
    });

    const fits = await t.copy(
      async (_ws, relay) => {
        await relay.putLink(
          id,
          "auth",
          new Uint8Array(atCap.blobBytes),
          atCap.widerRev ? WIDER_REV : 0,
        );
        return { label: "at cap" };
      },
      { copied: () => "copied", ready: () => "ready" },
    );
    expect(fits).toBe("copied");

    await expect(
      t.copy(
        async (_ws, relay) => {
          await relay.putLink(
            id,
            "auth",
            new Uint8Array(overCap.blobBytes),
            overCap.widerRev ? WIDER_REV : 0,
          );
          return { label: "over cap" };
        },
        { copied: () => "copied", ready: () => "ready" },
      ),
    ).rejects.toThrow(
      "That is too much to copy at once: copy one project or the selected links.",
    );
  });
});
