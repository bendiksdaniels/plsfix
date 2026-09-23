// @vitest-environment jsdom
// src/pane/links-transport.test.ts
// copy(): the clipboard write starts before run()'s first await and the
// bar/manual-box fallback shows exactly when the write does not land.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { BUNDLE_MAX_CHARS } from "../link/bundle";
import { WORKSPACE_STORAGE_KEY, type KeyStore } from "../link/workspace";
import type { Toast } from "../ui/toast";
import { installLinksTransport, type LinksTransport } from "./links-transport";

const ID = "a".repeat(32);

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

function fakeToast(): { toast: Toast; shown: string[] } {
  const shown: string[] = [];
  return { toast: { show: (message) => shown.push(message) }, shown };
}

async function transport(): Promise<LinksTransport> {
  return transportWith(memoryStore());
}

async function transportWith(
  keyStore: KeyStore,
  toast: Toast = fakeToast().toast,
): Promise<LinksTransport> {
  return installLinksTransport({ root: paneRoot(), keyStore, toast });
}

function copyReadyBar(): HTMLElement {
  return document.getElementById("copy-ready") as HTMLElement;
}

function transportSelect(): HTMLSelectElement {
  return document.getElementById("link-transport") as HTMLSelectElement;
}

function linkKeySection(): HTMLElement {
  return document.getElementById("link-key-section") as HTMLElement;
}

function autopushRow(): HTMLElement {
  return document.getElementById("links-autopush-row") as HTMLElement;
}

function pushSelectedLabel(): string {
  return document.getElementById("push-selected")?.textContent ?? "";
}

function pushAllLabel(): string {
  return document.getElementById("push-all")?.textContent ?? "";
}

function localHint(): HTMLElement {
  return document.getElementById("transport-local-hint") as HTMLElement;
}

function relayHint(): HTMLElement {
  return document.getElementById("transport-relay-hint") as HTMLElement;
}

function copyManual(): HTMLTextAreaElement {
  return document.getElementById("copy-manual") as HTMLTextAreaElement;
}

const lines = {
  copied: (r: { label: string }) => `Copied: ${r.label}`,
  ready: (r: { label: string }) => `Ready: ${r.label}`,
};

// A run() that records one link and one inbox row - enough for the bundle to
// count as non-empty, whatever the test wants to happen around it.
async function runOneLink(order?: string[]): Promise<{ label: string }> {
  order?.push("run");
  return { label: "Model!A1" };
}

function fillCollector(run: typeof runOneLink): (
  ws: unknown,
  relay: {
    putLink: (id: string, auth: string, blob: Uint8Array) => Promise<unknown>;
    postInbox: (
      ws: string,
      auth: string,
      id: string,
      blob: Uint8Array,
    ) => Promise<void>;
  },
) => Promise<{ label: string }> {
  return async (ws, relay) => {
    await relay.putLink(ID, "auth", new Uint8Array([1]));
    await relay.postInbox("ws", "auth", ID, new Uint8Array([2]));
    return run();
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("copy", () => {
  it("starts the clipboard write before run's first await and hides the bar on success", async () => {
    const order: string[] = [];
    const write = vi.fn(async () => undefined);
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", { clipboard: { write } });

    const t = await transport();
    const line = await t.copy(
      fillCollector(() => runOneLink(order)),
      lines,
    );

    expect(write).toHaveBeenCalledTimes(1);
    expect(order).toEqual(["run"]);
    expect(line).toBe("Copied: Model!A1");
    expect(copyReadyBar().hidden).toBe(true);
  });

  it("shows the ready line and bar when the write does not land", async () => {
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", {
      clipboard: { write: async () => Promise.reject(new Error("refused")) },
    });

    const t = await transport();
    const line = await t.copy(fillCollector(runOneLink), lines);

    expect(line).toBe("Ready: Model!A1");
    expect(copyReadyBar().hidden).toBe(false);
    expect(document.getElementById("copy-ready-text")?.textContent).toBe(
      "Ready: Model!A1",
    );
    expect(copyManual().hidden).toBe(true);
  });

  it("refuses a bundle over the character cap and rejects the pending write", async () => {
    const rejected: unknown[] = [];
    const write = vi.fn((items: FakeItem[]) => {
      const [html, plain] = Object.values(items[0]!.items);
      void html!.catch((error: unknown) => rejected.push(error));
      void plain!.catch((error: unknown) => rejected.push(error));
      return Promise.all([html, plain]);
    });
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", { clipboard: { write } });

    const t = await transport();
    const big = new Uint8Array(20 * 1024 * 1024);
    await expect(
      t.copy(async (_ws, relay) => {
        await relay.putLink(ID, "auth", big);
        return { label: "Big" };
      }, lines),
    ).rejects.toThrow(
      "That is too much to copy at once: copy one project or the selected links.",
    );
    await Promise.resolve();
    expect(rejected).toHaveLength(2);
    expect(copyReadyBar().hidden).toBe(true);
  });

  it("reports nothing to copy, without throwing, when the collector stayed empty", async () => {
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", {
      // The pending write rejects (nothing to copy): consume both flavor
      // promises here, the way a real clipboard.write() would, so the
      // rejection this test expects does not also surface as an unhandled
      // one.
      clipboard: {
        write: async (items: FakeItem[]) =>
          Promise.allSettled(Object.values(items[0]!.items)),
      },
    });

    const t = await transport();
    const line = await t.copy(async () => ({ label: "Model!A1" }), lines);
    expect(line).toBe("Nothing to copy: no link could be read.");
  });

  it("rejects the pending write and rethrows when run throws", async () => {
    const rejected: unknown[] = [];
    const write = vi.fn((items: FakeItem[]) => {
      const [html, plain] = Object.values(items[0]!.items);
      void html!.catch((error: unknown) => rejected.push(error));
      void plain!.catch((error: unknown) => rejected.push(error));
      return Promise.all([html, plain]);
    });
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", { clipboard: { write } });

    const t = await transport();
    await expect(
      t.copy(async () => {
        throw new Error("boom");
      }, lines),
    ).rejects.toThrow("boom");
    await Promise.resolve();
    expect(rejected).toHaveLength(2);
    expect(copyReadyBar().hidden).toBe(true);
  });
});

describe("retryCopy", () => {
  async function prepared(): Promise<LinksTransport> {
    vi.stubGlobal("ClipboardItem", FakeItem);
    vi.stubGlobal("navigator", {
      clipboard: { write: async () => Promise.reject(new Error("refused")) },
    });
    const t = await transport();
    await t.copy(fillCollector(runOneLink), lines);
    return t;
  }

  it("throws when nothing is prepared", async () => {
    const t = await transport();
    expect(() => t.retryCopy()).toThrow("Nothing is waiting to be copied.");
  });

  it("copies through execCommand and hides the bar on success", async () => {
    const t = await prepared();
    const data = new Map<string, string>();
    document.execCommand = vi.fn(() => {
      const event = new Event("copy", { cancelable: true });
      Object.defineProperty(event, "clipboardData", {
        value: { setData: (k: string, v: string) => data.set(k, v) },
      });
      document.dispatchEvent(event);
      return true;
    });

    expect(t.retryCopy()).toBe(
      "Copied for PowerPoint. Paste it in the pls,fix pane, Inbox tab.",
    );
    expect(data.get("text/plain")).toBeTruthy();
    expect(copyReadyBar().hidden).toBe(true);
    expect(() => t.retryCopy()).toThrow("Nothing is waiting to be copied.");
  });

  it("shows the manual box and throws when execCommand also fails", async () => {
    const t = await prepared();
    document.execCommand = vi.fn(() => false);

    expect(() => t.retryCopy()).toThrow(
      "Copy failed: select the text below and copy it by hand (Ctrl+C, or ⌘C on a Mac).",
    );
    const manual = copyManual();
    expect(manual.hidden).toBe(false);
    expect(manual.value).toContain('"plsfix"');
    expect(manual.value).toContain('"links"');
  });
});

describe("BUNDLE_MAX_CHARS", () => {
  it("is the 25 MB the spec names", () => {
    expect(BUNDLE_MAX_CHARS).toBe(25 * 1024 * 1024);
  });
});

describe("the default mode a boot reads", () => {
  it("is local when this device has never stored a link key", async () => {
    const t = await transportWith(memoryStore());
    expect(t.mode()).toBe("local");
  });

  it("is relay when this device already holds one", async () => {
    const store = memoryStore();
    await store.set(WORKSPACE_STORAGE_KEY, "k".repeat(43));
    const t = await transportWith(store);
    expect(t.mode()).toBe("relay");
  });

  it("applies before the tab's first render: the markup already matches", async () => {
    const store = memoryStore();
    await store.set(WORKSPACE_STORAGE_KEY, "k".repeat(43));
    await transportWith(store);
    expect(transportSelect().value).toBe("relay");
    expect(linkKeySection().hidden).toBe(false);
  });
});

describe("setMode: visibility and labels", () => {
  it("relay: shows the link key, auto-push and relay hint; hides the copy bar", async () => {
    const t = await transport();
    await t.setMode("relay");

    expect(t.mode()).toBe("relay");
    expect(transportSelect().value).toBe("relay");
    expect(linkKeySection().hidden).toBe(false);
    expect(autopushRow().hidden).toBe(false);
    expect(pushSelectedLabel()).toBe("Push selected");
    expect(pushAllLabel()).toBe("Push all");
    expect(localHint().hidden).toBe(true);
    expect(relayHint().hidden).toBe(false);
    expect(copyReadyBar().hidden).toBe(true);
  });

  it("local: hides the link key and auto-push, renames the push buttons", async () => {
    const store = memoryStore();
    await store.set(WORKSPACE_STORAGE_KEY, "k".repeat(43));
    const t = await transportWith(store); // boots relay
    await t.setMode("local");

    expect(t.mode()).toBe("local");
    expect(transportSelect().value).toBe("local");
    expect(linkKeySection().hidden).toBe(true);
    expect(autopushRow().hidden).toBe(true);
    expect(pushSelectedLabel()).toBe("Copy selected");
    expect(pushAllLabel()).toBe("Copy all");
    expect(localHint().hidden).toBe(false);
    expect(relayHint().hidden).toBe(true);
  });

  it("persists the choice for the next boot", async () => {
    const store = memoryStore();
    const t = await transportWith(store);
    await t.setMode("relay");
    expect(await store.get("plsfix.link.transport.v1")).toBe("relay");
    const rebooted = await transportWith(store);
    expect(rebooted.mode()).toBe("relay");
  });
});
