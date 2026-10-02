// @vitest-environment jsdom
// Attacks the two corners links-transport.test.ts does not drive:
// pendingLinksTransport() (what every call site meets before boot() has
// read the transport setting), the "Missing element" throw when the markup
// lacks a required id, and copy()/retryCopy() under three repeated presses.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { decodeBundle } from "../../src/link/bundle";
import type { KeyStore } from "../../src/link/workspace";
import type { Toast } from "../../src/ui/toast";
import {
  installLinksTransport,
  pendingLinksTransport,
  pushCopyLines,
} from "../../src/pane/links-transport";

class FakeItem {
  constructor(readonly items: Record<string, Promise<Blob>>) {}
}

function paneRoot(): Document {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  return document;
}

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

function fakeToast(): Toast {
  return { show: () => undefined };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("pendingLinksTransport: what a call meets before boot() finishes", () => {
  it("throws the same loading sentence from mode()", () => {
    const transport = pendingLinksTransport();
    expect(() => transport.mode()).toThrow(/Still starting up/);
  });

  it("throws it from retryCopy() too", () => {
    const transport = pendingLinksTransport();
    expect(() => transport.retryCopy()).toThrow(/Still starting up/);
  });

  it("rejects setMode() with the same sentence rather than throwing synchronously", async () => {
    const transport = pendingLinksTransport();
    await expect(transport.setMode("relay")).rejects.toThrow(
      /Still starting up/,
    );
  });

  it("rejects copy() with the same sentence, so a call site never has to null-check", async () => {
    const transport = pendingLinksTransport();
    await expect(
      transport.copy(async () => "unreachable", {
        copied: (r) => r,
        ready: (r) => r,
      }),
    ).rejects.toThrow(/Still starting up/);
  });

  it("keeps answering the same way on three repeated presses", () => {
    const transport = pendingLinksTransport();
    expect(() => transport.mode()).toThrow(/Still starting up/);
    expect(() => transport.mode()).toThrow(/Still starting up/);
    expect(() => transport.mode()).toThrow(/Still starting up/);
  });
});

describe("elements(): a pane whose markup lost a required id", () => {
  it("names the first missing id instead of a bare null-reference crash", async () => {
    const root = document.implementation.createHTMLDocument("");
    root.body.innerHTML = '<select id="link-transport"></select>';

    await expect(
      installLinksTransport({
        root,
        keyStore: memoryStore(),
        toast: fakeToast(),
      }),
    ).rejects.toThrow("Missing element #copy-ready");
  });
});

describe("copy(): three presses in a row while the clipboard keeps refusing", () => {
  // press distinguishes the pushed blob across repeated calls: "First" and
  // "Third" are both five characters, so keying the blob off label.length
  // alone would let a retryCopy that resent the FIRST bundle pass unnoticed.
  function fillOneLink(label: string, press = 0) {
    return async (
      _ws: unknown,
      relay: {
        putLink: (
          id: string,
          auth: string,
          blob: Uint8Array,
        ) => Promise<unknown>;
        postInbox: (
          ws: string,
          auth: string,
          id: string,
          blob: Uint8Array,
        ) => Promise<void>;
      },
    ) => {
      const id = "a".repeat(32);
      await relay.putLink(id, "auth", new Uint8Array([press]));
      await relay.postInbox("ws", "auth", id, new Uint8Array([1]));
      return { label };
    };
  }

  const lines = {
    copied: (r: { label: string }) => `Copied: ${r.label}`,
    ready: (r: { label: string }) => `Ready: ${r.label}`,
  };

  it("shows the ready line each time and keeps the manual box in step with the latest attempt, never a stale one", async () => {
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", {
      clipboard: { write: async () => Promise.reject(new Error("refused")) },
    });
    const transport = await installLinksTransport({
      root: paneRoot(),
      keyStore: memoryStore(),
      toast: fakeToast(),
    });

    for (const [index, label] of ["First", "Second", "Third"].entries()) {
      const line = await transport.copy(fillOneLink(label, index + 1), lines);
      expect(line).toBe(`Ready: ${label}`);
    }

    // retryCopy after three presses must retry the THIRD bundle, never the
    // first or second: decode the manual box's own bundle (the execCommand
    // fallback refused too, so it is the one thing left holding the prepared
    // JSON) and check the one link record it carries is press 3's, not
    // press 1's (both "First" and "Third" are five characters, so a
    // length-keyed blob could not tell them apart).
    document.execCommand = vi.fn(() => false);
    expect(() => transport.retryCopy()).toThrow(/Copy failed/);
    const manual = document.getElementById(
      "copy-manual",
    ) as HTMLTextAreaElement;
    const decoded = decodeBundle(manual.value);
    if (!decoded.ok) throw new Error("manual box did not hold a bundle");
    expect(decoded.bundle.links[0]?.blob).toEqual(new Uint8Array([3]));
  });

  it("never lets an old prepared bundle survive a newer, successful copy", async () => {
    vi.stubGlobal("ClipboardItem", FakeItem);
    const writes = vi.fn<() => Promise<void>>();
    vi.stubGlobal("navigator", {
      clipboard: {
        write: async () => {
          writes();
          return Promise.reject(new Error("refused"));
        },
      },
    });
    const transport = await installLinksTransport({
      root: paneRoot(),
      keyStore: memoryStore(),
      toast: fakeToast(),
    });

    await transport.copy(fillOneLink("First"), lines);
    expect(document.getElementById("copy-ready")?.hidden).toBe(false);

    vi.stubGlobal("navigator", {
      clipboard: { write: async () => undefined },
    });
    const line = await transport.copy(fillOneLink("Second"), lines);

    expect(line).toBe("Copied: Second");
    expect(document.getElementById("copy-ready")?.hidden).toBe(true);
    // A third press with the clipboard failing again must show ONLY the
    // third bundle, not resurrect the first one hideBar() already cleared.
    vi.stubGlobal("navigator", {
      clipboard: { write: async () => Promise.reject(new Error("refused")) },
    });
    const third = await transport.copy(fillOneLink("Third"), lines);
    expect(third).toBe("Ready: Third");
    expect(document.getElementById("copy-ready-text")?.textContent).toBe(
      "Ready: Third",
    );
  });
});

describe("pushCopyLines: the ready line, not only the copied one", () => {
  it("says the push counts as copies before the clipboard write is confirmed", () => {
    expect(
      pushCopyLines().ready({
        pushed: 2,
        missing: 1,
        failed: 0,
        failures: [],
      }),
    ).toBe("2 copied, 1 missing, 0 failed: press Copy for PowerPoint.");
  });
});
