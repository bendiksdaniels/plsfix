// @vitest-environment jsdom
// Uncovered edges of the shortcut manager: installing on a root with no
// #shortcuts-list section at all (the doc comment's "a pane without the
// section is left alone"), a requirement check that throws instead of
// answering false (a host mid-boot), a non-ok HTTP status fetching
// shortcuts.json, the fetch's own 5 s deadline actually firing, and a box
// that has gone missing from the DOM by the time a read tries to fill it.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  installOffice,
  installPanel,
  makeActions,
  note,
  paneRoot,
  rowCount,
  serveShortcuts,
  box,
} from "../../test/shortcuts.support";

beforeEach(() => {
  vi.resetModules();
  paneRoot();
  serveShortcuts();
});

describe("installShortcutsPanel on a root with no section", () => {
  it("returns without fetching or touching anything (e.g. the deck's pane)", async () => {
    const fetchSpy = vi.fn();
    Object.assign(globalThis, { fetch: fetchSpy });
    installOffice(makeActions());
    const module = await import("../../src/pane/shortcuts-panel");

    const bareRoot = document.createElement("div"); // no #shortcuts-list
    const onCard = vi.fn();
    await expect(
      module.installShortcutsPanel(bareRoot, onCard),
    ).resolves.toBeUndefined();
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("still renders when the section exists with no tabpanel ancestor", async () => {
    installOffice(makeActions());
    const module = await import("../../src/pane/shortcuts-panel");

    const root = document.createElement("div");
    root.innerHTML =
      '<div id="shortcuts-list"></div><p id="shortcuts-note"></p>';
    await module.installShortcutsPanel(root, vi.fn());

    expect(root.querySelectorAll("#shortcuts-list label").length).toBe(42);
  });
});

describe("a requirement check that throws instead of answering false", () => {
  it("is treated as unsupported, not a crash", async () => {
    Object.assign(globalThis, {
      Office: {
        context: {
          requirements: {
            isSetSupported: () => {
              throw new Error("mid-boot: requirements not ready");
            },
          },
        },
        actions: makeActions(),
      },
    });

    await installPanel();
    expect(rowCount()).toBe(42);
    expect(note()).toBe(
      "Custom shortcuts need Microsoft 365 with a signed-in account.",
    );
    expect(box("PLSFIX_AUTOCOLOR").disabled).toBe(true);
  });
});

describe("fetching public/shortcuts.json", () => {
  it("treats a non-ok HTTP status the same as an offline fetch", async () => {
    Object.assign(globalThis, {
      fetch: vi.fn(async () => ({
        ok: false,
        status: 404,
        json: async () => ({}) as unknown,
      })),
    });
    installOffice(makeActions());

    await installPanel();
    expect(note()).toContain("could not be read");
    expect(rowCount()).toBe(0);
  });

  it("aborts a fetch that never answers once its own deadline passes", async () => {
    // Imported (and its own dynamic-import microtasks settled) BEFORE fake
    // timers go on: installPanel()'s own `await import(...)` should not
    // race the clock this test is about to take over.
    const module = await import("../../src/pane/shortcuts-panel");
    // A real fetch rejects with an AbortError once its signal fires; this
    // stub is abort-aware for exactly that reason (a bare hanging promise
    // would never notice the controller.abort() call at all).
    Object.assign(globalThis, {
      fetch: vi.fn(
        (_url: string, options?: { signal?: AbortSignal }) =>
          new Promise((_resolve, reject) => {
            options?.signal?.addEventListener("abort", () => {
              const error = new Error("aborted");
              error.name = "AbortError";
              reject(error);
            });
          }),
      ),
    });
    installOffice(makeActions());

    vi.useFakeTimers();
    try {
      const installing = module.installShortcutsPanel(document, vi.fn());
      await vi.advanceTimersByTimeAsync(5_000); // FETCH_TIMEOUT_MS
      await installing;

      expect(note()).toContain("could not be read");
      expect(rowCount()).toBe(0);
    } finally {
      vi.useRealTimers();
    }
  });
});

describe("a box that has gone missing from the DOM", () => {
  it("skips it on the next read instead of throwing", async () => {
    installOffice(makeActions());
    const { applyShortcuts } = await installPanel();
    box("PLSFIX_AUTOCOLOR").remove();

    // applyShortcuts' final fill(api, true) re-reads every row; the one
    // whose box is gone must be skipped, not thrown over, and every other
    // box must still get filled.
    await expect(applyShortcuts()).resolves.toContain("updated");
    expect(document.getElementById("shortcut-PLSFIX_AUTOCOLOR")).toBeNull();
    expect(box("PLSFIX_SHOWPANE").disabled).toBe(false);
  });
});
