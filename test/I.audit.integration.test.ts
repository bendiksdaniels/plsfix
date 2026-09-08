// @vitest-environment jsdom
// Slice I audit: src/main.ts has never run in a test (0% coverage). Boots
// the real taskpane.html + src/main.ts over test/fakehost.ts's fake Excel
// host in the three shapes boot() actually branches on - a fully connected
// host, an ExcelApi too old to use, and a host that answers Office.onReady
// with something other than Excel or PowerPoint (the closest a jsdom test
// gets to "no Office host at all"; a real headless-Chromium run of this
// exact scenario, driven through office.js from the CDN, is in the ux:sweep
// gate). Every shape must still let the pane switch tabs, open help, search
// and answer a click - never a dead button, never an uncaught rejection.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import type * as AuditModule from "../src/excel/audit";
import { restorePersistedOverlay } from "../src/excel/audit";
import { installFakeHost, uninstallFakeHost } from "./fakehost";

// Wraps the real restorePersistedOverlay so every test but one calls straight
// through to it (against the fake host installed below); only the "B's
// finding" test overrides it for one call, to force the one failure it needs
// without racing every other fire-and-forget Excel.run boot() also starts.
vi.mock("../src/excel/audit", async (importOriginal) => {
  const actual = await importOriginal<typeof AuditModule>();
  return {
    ...actual,
    restorePersistedOverlay: vi.fn(actual.restorePersistedOverlay),
  };
});

function pane(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

interface BootOptions {
  isSetSupported?: (set: string, version: string) => boolean;
  /** Office.onReady answers with this host name instead of "Excel". */
  hostOverride?: string;
}

// test/fakehost.ts's Office.onReady stub only implements the callback style
// (`callback?.({host: "Excel"})` resolves to undefined with no callback);
// src/main.ts calls the no-callback, promise-returning style, which real
// office.js also supports. Nothing here has ever driven main.ts/hostReady()
// through the fake host before (0% coverage), so the gap was never hit -
// worked around locally rather than touched in the shared file (fakehost.ts
// edits are additive-only, and this needs an existing line changed).
function patchOnReadyForPromiseStyle(hostOverride?: string): void {
  const office = (
    globalThis as unknown as {
      Office: {
        onReady: (cb?: (info: { host: string }) => unknown) => unknown;
      };
    }
  ).Office;
  office.onReady = (callback) => {
    const info = { host: hostOverride ?? "Excel" };
    callback?.(info);
    return Promise.resolve(info);
  };
}

// The pane boots on import (main.ts wires everything from its own top-level
// code), so the markup and the fake host go in first, exactly like
// test/ppt.boot.integration.test.ts does for the PowerPoint pane.
async function boot(options: BootOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  pane();
  installFakeHost({ isSetSupported: options.isSetSupported });
  patchOnReadyForPromiseStyle(options.hostOverride);
  await import("../src/main");
  // hostReady's headStartMs/give-up timers are real (host-ready.ts is not
  // faked here), but the "ready" race settles on a microtask once
  // Office.onReady resolves - a couple of ticks is enough to reach boot().
  await new Promise((resolve) => setTimeout(resolve, 0));
  await new Promise((resolve) => setTimeout(resolve, 0));
}

// Most action-list buttons carry only data-action, not an id; a few (like
// shortcut-card) carry both. Either lookup finds a real click target.
function click(idOrAction: string): void {
  const button =
    document.getElementById(idOrAction) ??
    document.querySelector<HTMLButtonElement>(`[data-action="${idOrAction}"]`);
  if (!button) throw new Error(`test setup: no button for "${idOrAction}"`);
  button.click();
}

async function settle(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

function badge(): string {
  return document.getElementById("connection-status")?.textContent ?? "";
}

function search(query: string): void {
  const input = document.getElementById("tool-search") as HTMLInputElement;
  input.value = query;
  input.dispatchEvent(new Event("input"));
}

afterEach(() => {
  uninstallFakeHost();
});

describe("independence: no usable Excel host", () => {
  it.each([
    ["a host that is not Excel at all", { hostOverride: "Word" }],
    ["ExcelApi too old for the pane", { isSetSupported: () => false }],
  ] as const)(
    "still switches tabs and opens help - %s",
    async (_label, opts) => {
      await boot(opts);

      click("tab-links");
      expect(document.getElementById("view-links")!.hidden).toBe(false);
      expect(document.getElementById("view-tools")!.hidden).toBe(true);

      click("tab-tools");
      expect(document.getElementById("view-tools")!.hidden).toBe(false);

      const helpToggle = document
        .getElementById("selection-heading")!
        .closest(".section-heading")!
        .querySelector<HTMLButtonElement>(".help-toggle");
      expect(helpToggle).not.toBeNull();
      helpToggle!.click();
      expect(helpToggle!.getAttribute("aria-expanded")).toBe("true");
    },
  );

  it.each([
    ["a host that is not Excel at all", { hostOverride: "Word" }],
    ["ExcelApi too old for the pane", { isSetSupported: () => false }],
  ] as const)("still runs the search box - %s", async (_label, opts) => {
    await boot(opts);
    search("waterfall");
    const results = document.getElementById("tool-search-results")!;
    expect(results.hidden).toBe(false);
    expect(results.textContent).toContain("Waterfall");
  });

  it.each([
    ["a host that is not Excel at all", { hostOverride: "Word" }],
    ["ExcelApi too old for the pane", { isSetSupported: () => false }],
  ] as const)(
    "refuses a workbook action with one sentence, not a crash - %s",
    async (_label, opts) => {
      await boot(opts);
      expect(() => click("undo")).not.toThrow();
      await settle();
      expect(toastText()).toBe("Excel is not connected.");
      expect(document.getElementById("toast")?.className).toContain("error");
    },
  );

  it("shows the right badge for each rejected shape", async () => {
    await boot({ hostOverride: "Word" });
    expect(badge()).toBe("Excel required");

    await boot({ isSetSupported: () => false });
    expect(badge()).toBe("Excel 2021 / Microsoft 365 required");
  });

  it("still opens the shortcut card - Office chrome, not a workbook action", async () => {
    // The fake host's Office.context.ui has no displayDialogAsync, so this
    // exercises openShortcutCard's window.open fallback; jsdom does not
    // implement navigation, so the fallback itself is stubbed out here.
    vi.stubGlobal("open", vi.fn());
    await boot({ hostOverride: "Word" });
    click("shortcut-card");
    await settle();
    expect(toastText()).toBe("Shortcut card opened");
    vi.unstubAllGlobals();
  });
});

describe("a fully connected Excel host", () => {
  it("connects, wires the search and runs a real action end to end", async () => {
    await boot();
    expect(badge()).toBe("Excel connected");

    search("waterfall");
    const results = document.getElementById("tool-search-results")!;
    expect(results.hidden).toBe(false);

    click("undo");
    await settle();
    expect(toastText()).toBe("There is no pls,fix action to undo yet.");
  });
});

describe("restoreOverlayFills (B's finding: a silent catch left no trace)", () => {
  it("toasts one sentence when the overlay restore fails instead of staying silent", async () => {
    vi.mocked(restorePersistedOverlay).mockRejectedValueOnce(
      new Error("workbook settings refused"),
    );
    await boot();

    expect(toastText()).toBe(
      "Last session's audit overlay could not be put back: toggle it once to clear the stripes.",
    );
    expect(document.getElementById("toast")?.className).toContain("error");
  });

  it("still restores the overlay and toasts normally when it succeeds", async () => {
    vi.mocked(restorePersistedOverlay).mockResolvedValueOnce(true);
    await boot();

    expect(toastText()).toBe(
      "Audit overlay fills from the last session were restored.",
    );
    expect(document.getElementById("toast")?.className).not.toContain("error");
  });

  it("stays quiet when there was nothing to restore", async () => {
    vi.mocked(restorePersistedOverlay).mockResolvedValueOnce(false);
    await boot();

    expect(toastText()).toBe("");
  });
});
