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

// links-tab.ts's own installLinksTab() fires a further, un-awaited boot chain
// of its own (key load, list refresh, relay touch, toggle restore - each a
// separate await), fully independent of anything main.ts's caller waits on.
// A couple of ticks settles hostReady()/boot() itself, but leaves that inner
// chain mid-flight; the next test's uninstallFakeHost() then pulls Office/
// Excel out from under it, surfacing as an unrelated unhandled rejection.
// Ten rounds of a macrotask each drains every reasonable depth of chained
// promises without a real-time sleep.
async function drain(rounds = 10): Promise<void> {
  for (let i = 0; i < rounds; i += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
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
  await drain();
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
  await drain();
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

  // Both delete confirms are wired directly in main.ts (not through
  // dispatch()'s data-action loop): the first click only arms the button,
  // the second runs the real delete through the guard.
  it("arms delete-names on the first click, runs it on the second", async () => {
    await boot();
    click("delete-names");
    await settle();
    expect(
      document.getElementById("delete-names")?.classList.contains("armed"),
    ).toBe(true);

    click("delete-names");
    await settle();
    expect(toastText()).toMatch(/^Deleted \d+ broken names?$/);
    expect(
      document.getElementById("delete-names")?.classList.contains("armed"),
    ).toBe(false);
  });

  it("arms styles-delete on the first click, runs it on the second", async () => {
    await boot();
    click("styles-delete");
    await settle();
    expect(
      document.getElementById("styles-delete")?.classList.contains("armed"),
    ).toBe(true);

    click("styles-delete");
    await settle();
    expect(toastText()).toMatch(/^Deleted \d+ unused styles?$/);
    expect(
      document.getElementById("styles-delete")?.classList.contains("armed"),
    ).toBe(false);
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

describe("ribbon commands (H's finding: registered too late)", () => {
  // Bypasses the shared boot() helper: this checks state immediately after
  // import, before any settle tick - the whole point is that registration
  // must not wait on hostReady()/boot() to reach a connected Excel.
  async function importFresh(): Promise<{
    actions: Map<string, (event?: { completed: () => void }) => void>;
  }> {
    vi.resetModules();
    uninstallFakeHost();
    pane();
    const { helpers } = installFakeHost();
    patchOnReadyForPromiseStyle();
    await import("../src/main");
    return { actions: helpers.actions() };
  }

  it("associates every PLSFIX_* command before the host probe ever settles", async () => {
    // A quick-resolving Office.onReady is not a strong enough test here: the
    // fake's whole hostReady -> boot -> connectExcel chain is plain
    // Promise.resolve()s with no real macrotask boundary, so it can finish
    // inside the same dynamic import() the assertion runs after regardless
    // of where registerCommands() is called from. A promise this test
    // resolves itself, once the assertion is already done, is the only way
    // to prove registration did not wait on it.
    vi.resetModules();
    uninstallFakeHost();
    pane();
    const { helpers } = installFakeHost();
    let resolveOnReady!: (info: { host: string }) => void;
    const office = (
      globalThis as unknown as {
        Office: { onReady: () => Promise<{ host: string }> };
      }
    ).Office;
    office.onReady = () => new Promise((resolve) => (resolveOnReady = resolve));

    await import("../src/main");

    expect(helpers.actions().has("PLSFIX_UNDO")).toBe(true);
    expect(helpers.actions().has("PLSFIX_AUTOCOLOR")).toBe(true);
    expect(helpers.actions().has("PLSFIX_SHOWPANE")).toBe(true);

    // Let boot() proceed and fully settle so nothing is left running for the
    // next test's fresh (or absent) fake host to trip over.
    resolveOnReady({ host: "Excel" });
    await drain();
  });

  it("a command fired before Excel connects reports through its own promise chain, not a crash", async () => {
    const { actions } = await importFresh();
    const handler = actions.get("PLSFIX_UNDO");
    expect(handler).toBeDefined();

    let completed = false;
    expect(() =>
      handler!({ completed: () => (completed = true) }),
    ).not.toThrow();
    await drain();

    expect(completed).toBe(true);
    expect(toastText()).toBe("There is no pls,fix action to undo yet.");
  });
});

describe("PowerPoint host: redirects instead of booting the Excel pane", () => {
  it("replaces the page with pptpane.html", async () => {
    // jsdom's own location.replace is not configurable in place (navigation
    // is not implemented), so the whole global is swapped for a plain stub.
    const replaceSpy = vi.fn();
    vi.stubGlobal("location", { ...window.location, replace: replaceSpy });
    await boot({ hostOverride: "PowerPoint" });
    expect(replaceSpy).toHaveBeenCalledWith("pptpane.html");
    vi.unstubAllGlobals();
  });
});

describe("degraded boot: Office.onReady never settles but Excel answers a real probe", () => {
  it("boots connected via the probe, with the degraded-boot toast", async () => {
    vi.resetModules();
    uninstallFakeHost();
    pane();
    installFakeHost();
    const office = (
      globalThis as unknown as {
        Office: {
          onReady: () => Promise<never>;
          context: { host?: string };
          HostType: { Excel: string };
        };
      }
    ).Office;
    // Never settles: the exact case src/host-ready.ts's probe exists for
    // (Excel for the web, custom-functions init stuck) - Excel.run still
    // answers underneath it, once asked. test/fakehost.ts's Office.context
    // carries no host field at all (nothing has driven probeExcelHost()
    // through it before), so it is set here too, the same local-workaround
    // way as the onReady style above.
    office.onReady = () => new Promise(() => undefined);
    office.context.host = office.HostType.Excel;

    vi.useFakeTimers();
    await import("../src/main");
    // HOST_HEAD_START_MS (4000) then the first probe attempt.
    await vi.advanceTimersByTimeAsync(4_100);
    vi.useRealTimers();
    await drain();

    expect(document.getElementById("connection-status")?.textContent).toBe(
      "Excel connected",
    );
    expect(toastText()).toBe(
      "Excel answered but never reported the add-in ready: =PLSFIX.ROUND may need the workbook reopened.",
    );
  });

  it("retries after a probe attempt that genuinely fails", async () => {
    vi.resetModules();
    uninstallFakeHost();
    pane();
    const { helpers } = installFakeHost();
    const office = (
      globalThis as unknown as {
        Office: {
          onReady: () => Promise<never>;
          context: { host?: string };
          HostType: { Excel: string };
        };
      }
    ).Office;
    office.onReady = () => new Promise(() => undefined);
    office.context.host = office.HostType.Excel;
    // The first probe's own Excel.run rejects for a real reason (not "no
    // host") - probeExcelHost's catch must still answer null and let
    // src/host-ready.ts's loop try again, rather than getting stuck.
    helpers.failNextSync();

    vi.useFakeTimers();
    await import("../src/main");
    await vi.advanceTimersByTimeAsync(4_100);
    expect(document.getElementById("connection-status")?.textContent).toBe(
      "Connecting",
    );
    // HOST_PROBE_EVERY_MS (1000) to the second, unobstructed attempt.
    await vi.advanceTimersByTimeAsync(1_100);
    vi.useRealTimers();
    await drain();

    expect(document.getElementById("connection-status")?.textContent).toBe(
      "Excel connected",
    );
  });
});
