// @vitest-environment jsdom
// Attacks the clipboard write's refusal chain (src/ui/clipboard-links.ts)
// past what clipboard-links.test.ts covers: clipboard.write() throwing
// SYNCHRONOUSLY, a "copy" event with no clipboardData, and that
// beginClipboardWrite leaves no unhandled rejection with no ClipboardItem at all.

import { afterEach, describe, expect, it, vi } from "vitest";
import {
  BUNDLE_SENTENCE,
  bundleHtml,
  encodeBundle,
  type Bundle,
} from "../../src/link/bundle";
import {
  beginClipboardWrite,
  bundleFromPaste,
  copyBundleNow,
} from "../../src/ui/clipboard-links";

afterEach(() => {
  vi.unstubAllGlobals();
});

class FakeItem {
  constructor(readonly items: Record<string, Promise<Blob>>) {}
}

describe("beginClipboardWrite: clipboard.write() throwing synchronously", () => {
  it("answers false instead of throwing out of beginClipboardWrite itself", async () => {
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", {
      clipboard: {
        write: () => {
          // Some WebKit builds throw here directly rather than returning a
          // rejected promise - the call itself is where NotAllowedError
          // lands, not the promise clipboard.write() would otherwise give.
          throw new DOMException("Document is not focused.", "NotAllowedError");
        },
      },
    });

    const pending = beginClipboardWrite();
    pending.resolve("{}");

    await expect(pending.done).resolves.toBe(false);
  });

  it("leaves the caller free to fall back to Copy for PowerPoint afterwards", async () => {
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", {
      clipboard: {
        write: () => {
          throw new DOMException("no", "NotAllowedError");
        },
      },
    });

    const pending = beginClipboardWrite();
    pending.resolve("{}");
    expect(await pending.done).toBe(false);

    // The synchronous execCommand fallback still works on the same content:
    // beginClipboardWrite's own throw must not have consumed or corrupted it.
    document.execCommand = vi.fn(() => {
      const event = new Event("copy", { cancelable: true });
      Object.defineProperty(event, "clipboardData", {
        value: { setData: () => undefined },
      });
      document.dispatchEvent(event);
      return true;
    });
    expect(copyBundleNow("{}")).toBe(true);
  });
});

describe("copyBundleNow: a copy event with no clipboardData", () => {
  it("answers false without throwing when the engine never populates clipboardData", () => {
    document.execCommand = vi.fn(() => {
      const event = new Event("copy", { cancelable: true });
      // No clipboardData property at all: some hosts fire the event but
      // never attach the DataTransfer the handler needs.
      document.dispatchEvent(event);
      return true;
    });

    expect(copyBundleNow("{}")).toBe(false);
  });

  it("answers false when execCommand itself throws", () => {
    document.execCommand = vi.fn(() => {
      throw new Error("document.execCommand is not supported");
    });

    expect(copyBundleNow("{}")).toBe(false);
  });

  it("removes its own copy listener even after refusing, so a later real copy is not doubled", () => {
    const seen: string[] = [];
    const outside = (): number => seen.push("outside");
    document.addEventListener("copy", outside);
    try {
      document.execCommand = vi.fn(() => {
        const event = new Event("copy", { cancelable: true });
        document.dispatchEvent(event); // no clipboardData: copyBundleNow's own handler no-ops
        return true;
      });

      copyBundleNow("{}");

      // This second dispatch DOES carry clipboardData, unlike the first: a
      // leaked listener (the finally's removeEventListener deleted) would
      // call setData on it. Only the outside listener may still be attached.
      const setData = vi.fn();
      const second = new Event("copy", { cancelable: true });
      Object.defineProperty(second, "clipboardData", { value: { setData } });
      document.dispatchEvent(second);

      // The outside listener heard both dispatches; copyBundleNow's own
      // listener must be gone after the first, or a second unrelated copy
      // event later in the session would silently write the stale bundle.
      expect(seen).toEqual(["outside", "outside"]);
      expect(setData).not.toHaveBeenCalled();
    } finally {
      document.removeEventListener("copy", outside);
    }
  });
});

describe("beginClipboardWrite: no unhandled rejection when the engine has no ClipboardItem", () => {
  it("settles quietly even when the caller rejects the pending write", async () => {
    vi.stubGlobal("ClipboardItem", undefined);
    const rejections: unknown[] = [];
    const onRejection = (reason: unknown): void => {
      rejections.push(reason);
    };
    process.on("unhandledRejection", onRejection);
    try {
      const pending = beginClipboardWrite();
      pending.reject(new Error("nothing to copy"));
      expect(await pending.done).toBe(false);
      // Give any unhandled rejection a turn to surface before asserting none did.
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(rejections).toEqual([]);
    } finally {
      process.off("unhandledRejection", onRejection);
    }
  });
});

describe("bundleFromPaste: an actual paste, not just the null guard", () => {
  const ID = "0123456789abcdef0123456789abcdef";
  const sample: Bundle = {
    links: [
      {
        id: ID,
        rev: 2 ** 40 + 1,
        sentAt: 1_790_000_000_000,
        blob: new Uint8Array([1, 2, 3]),
      },
    ],
    inbox: [],
  };

  function dataTransfer(html: string, plain: string): DataTransfer {
    const values: Record<string, string> = {
      "text/html": html,
      "text/plain": plain,
    };
    return {
      getData: (type: string) => values[type] ?? "",
    } as unknown as DataTransfer;
  }

  it("reads the HTML flavor of a real paste event's DataTransfer", () => {
    const html = bundleHtml(encodeBundle(sample));
    const data = dataTransfer(html, BUNDLE_SENTENCE);

    expect(bundleFromPaste(data)).toEqual({ ok: true, bundle: sample });
  });

  it("falls back to the plain flavor when there is no carrier in the html", () => {
    const data = dataTransfer("", encodeBundle(sample));
    expect(bundleFromPaste(data)).toEqual({ ok: true, bundle: sample });
  });

  it("answers notBundle for a paste of ordinary text, not [object Object]", () => {
    const data = dataTransfer("<b>hello</b>", "hello");
    expect(bundleFromPaste(data)).toEqual({ ok: false, reason: "notBundle" });
  });
});
