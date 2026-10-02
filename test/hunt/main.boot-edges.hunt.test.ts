// @vitest-environment jsdom
// Uncovered edges of src/main.ts's boot(): the top-level error-reporting wire
// (a real window "error"/"unhandledrejection" reaching the toast), the
// find-query's Enter-key shortcut (never pressed by the clickthrough sweep,
// which only ever presses the "find" button), and the debounced
// DocumentSelectionChanged handler (repeats lens: three rapid fires must
// settle into exactly one refresh of the LATEST selection, not a stack of
// three, and not a crash).

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "../fakehost";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let selectionHandler: (() => void) | null = null;

function pane(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

// Same technique test/clickthrough.excel.integration.test.ts uses for
// Office.onReady's promise style; the fake's own addHandlerAsync is a no-op
// that never remembers the callback, so a selection-changed event cannot be
// simulated without capturing it here first.
function patchOfficeForBoot(): void {
  const office = (
    globalThis as unknown as {
      Office: {
        onReady: (cb?: (info: { host: string }) => unknown) => unknown;
        context: {
          document: {
            addHandlerAsync: (type: unknown, handler: () => void) => void;
          };
        };
      };
    }
  ).Office;
  office.onReady = (callback) => {
    const info = { host: "Excel" };
    callback?.(info);
    return Promise.resolve(info);
  };
  office.context.document.addHandlerAsync = (_type, handler) => {
    selectionHandler = handler;
  };
}

async function drain(rounds = 12): Promise<void> {
  for (let i = 0; i < rounds; i += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

function paneBusy(): boolean {
  const tabs = document.querySelectorAll<HTMLButtonElement>("[role=tab]");
  return tabs.length > 0 && [...tabs].every((tab) => tab.disabled);
}

async function settle(): Promise<void> {
  await drain();
  for (let i = 0; i < 200 && paneBusy(); i += 1) await drain(1);
}

vi.stubGlobal("open", vi.fn());

async function bootExcel(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  selectionHandler = null;
  pane();
  helpers = installFakeHost({ sheets: ["Model", "Data"] }).helpers;
  patchOfficeForBoot();
  await import("../../src/main");
  await settle();
}

afterEach(() => {
  uninstallFakeHost();
  vi.restoreAllMocks();
  vi.stubGlobal("open", vi.fn());
});

describe("main.ts boot edges", () => {
  it("wires a real window error to the toast through installErrorReporting", async () => {
    await bootExcel();
    const before = document.querySelector("#toast .toast-text")?.textContent;

    const event = new ErrorEvent("error", { error: new Error("boom") });
    window.dispatchEvent(event);
    await settle();

    const after = document.querySelector("#toast .toast-text")?.textContent;
    expect(after).not.toBe(before);
    // describeError(error, ctx) with no action: a real Error's own message
    // travels unchanged, only a non-Error thrown value gets FALLBACK_MESSAGE.
    expect(after).toBe("boom");
  });

  it("wires a real unhandled rejection to the toast too", async () => {
    await bootExcel();
    // A real rejected promise, caught right away so Node's own
    // unhandledRejection detection never sees it as unhandled - this test
    // is only about report.ts's DOM "unhandledrejection" listener, fired
    // here by hand since jsdom does not implement PromiseRejectionEvent
    // and never dispatches this event on its own for a promise nothing
    // else observes.
    const settled = Promise.reject(new Error("rejected"));
    settled.catch(() => undefined);
    const event = new Event("unhandledrejection") as Event & {
      reason?: unknown;
      promise?: Promise<unknown>;
    };
    Object.defineProperty(event, "reason", { value: new Error("rejected") });
    Object.defineProperty(event, "promise", { value: settled });
    window.dispatchEvent(event);
    await settle();

    expect(document.querySelector("#toast .toast-text")?.textContent).toBe(
      "rejected",
    );
  });

  it("runs Find on Enter in the query box, the same as the Find button", async () => {
    await bootExcel();
    helpers.seed("Model!A1", [["Revenue"]]);
    const findTab = document.getElementById("tab-tools") as HTMLButtonElement;
    findTab?.click();
    await settle();

    const input = document.getElementById("find-query") as HTMLInputElement;
    const hintBefore = document.getElementById("find-hint")?.textContent;
    input.value = "Revenue";
    input.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    await settle();

    const hintAfter = document.getElementById("find-hint")?.textContent;
    expect(hintAfter).not.toBe(hintBefore);
    expect(hintAfter).toMatch(/hit|Searched/i);
  });

  it("a non-Enter key on the query box does nothing", async () => {
    await bootExcel();
    const findTab = document.getElementById("tab-tools") as HTMLButtonElement;
    findTab?.click();
    await settle();

    const input = document.getElementById("find-query") as HTMLInputElement;
    const hintBefore = document.getElementById("find-hint")?.textContent;
    input.value = "Revenue";
    input.dispatchEvent(
      new KeyboardEvent("keydown", { key: "a", bubbles: true }),
    );
    await settle();
    expect(document.getElementById("find-hint")?.textContent).toBe(hintBefore);
  });

  describe("debounced selection-changed", () => {
    it("collapses three rapid fires into one refresh of the LATEST selection", async () => {
      await bootExcel();
      expect(selectionHandler).not.toBeNull();

      helpers.select("Model!A1");
      selectionHandler!();
      helpers.select("Model!B2:C2");
      selectionHandler!();
      helpers.select("Model!D4:D6");
      selectionHandler!(); // only this last one should ever land

      // Real timers: the debounce is a real window.setTimeout(150ms), and
      // this file's settle() only yields microtasks/macrotask ticks, not
      // wall-clock time.
      await new Promise((resolve) => setTimeout(resolve, 250));
      await settle();

      expect(document.getElementById("selection-address")?.textContent).toBe(
        "Model!D4:D6",
      );
    });

    it("does nothing until the debounce window elapses", async () => {
      await bootExcel();
      const before = document.getElementById("selection-address")?.textContent;
      helpers.select("Model!F9");
      selectionHandler!();
      // No wait at all: the 150 ms window has not elapsed yet.
      expect(document.getElementById("selection-address")?.textContent).toBe(
        before,
      );
      await new Promise((resolve) => setTimeout(resolve, 250));
      expect(document.getElementById("selection-address")?.textContent).toBe(
        "Model!F9",
      );
    });
  });
});
