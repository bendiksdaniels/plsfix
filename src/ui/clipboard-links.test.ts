// @vitest-environment jsdom
// src/ui/clipboard-links.test.ts
// The write starts inside the press and its content arrives later; the
// synchronous fallback sets both flavors; a paste reads the bundle back.
import { afterEach, describe, expect, it, vi } from "vitest";
import { BUNDLE_SENTENCE, bundleHtml } from "../link/bundle";
import {
  beginClipboardWrite,
  bundleFromPaste,
  copyBundleNow,
} from "./clipboard-links";

class FakeItem {
  constructor(readonly items: Record<string, Promise<Blob>>) {}
}

// jsdom 26's Blob has no .text() (only slice/size/type - a known jsdom gap),
// so the fake reads it back through FileReader instead, which jsdom does
// implement correctly. A real webview's Blob has .text(); this is test-only.
function readBlobText(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error as unknown);
    reader.readAsText(blob);
  });
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("beginClipboardWrite", () => {
  it("calls write at once and lands both flavors when the content arrives", async () => {
    const write = vi.fn(async (items: FakeItem[]) => {
      const blobs = await Promise.all(Object.values(items[0]!.items));
      expect(await readBlobText(blobs[0]!)).toBe(bundleHtml("{}"));
      expect(await readBlobText(blobs[1]!)).toBe(BUNDLE_SENTENCE);
    });
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", { clipboard: { write } });
    const pending = beginClipboardWrite();
    expect(write).toHaveBeenCalledTimes(1);
    pending.resolve("{}");
    await expect(pending.done).resolves.toBe(true);
  });

  it("answers false when the engine has no ClipboardItem or refuses", async () => {
    vi.stubGlobal("ClipboardItem", undefined);
    await expect(beginClipboardWrite().done).resolves.toBe(false);
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", {
      clipboard: {
        write: async () => {
          throw new DOMException("no", "NotAllowedError");
        },
      },
    });
    const pending = beginClipboardWrite();
    pending.resolve("{}");
    await expect(pending.done).resolves.toBe(false);
  });
});

describe("copyBundleNow", () => {
  it("sets both flavors through a copy event", () => {
    const data = new Map<string, string>();
    document.execCommand = vi.fn(() => {
      const event = new Event("copy", { cancelable: true });
      Object.defineProperty(event, "clipboardData", {
        value: { setData: (t: string, v: string) => data.set(t, v) },
      });
      document.dispatchEvent(event);
      return true;
    });
    expect(copyBundleNow("{}")).toBe(true);
    expect(data.get("text/html")).toBe(bundleHtml("{}"));
    expect(data.get("text/plain")).toBe(BUNDLE_SENTENCE);
  });
});

describe("bundleFromPaste", () => {
  it("reads nothing from nothing", () => {
    expect(bundleFromPaste(null)).toEqual({ ok: false, reason: "notBundle" });
  });
});
