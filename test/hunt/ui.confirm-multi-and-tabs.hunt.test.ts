// @vitest-environment jsdom
// Two edges of the two-click confirm named in the brief that confirm.test.ts
// does not cover: a second button armed while the first is still armed (each
// instance owns its own timer and label, arming one must never touch the
// other), and a tab switch away and back while a button sits armed (confirm.ts
// is tab-agnostic by design - nothing disarms a button just because its panel
// went out of view, so it must still be armed, and still require the second
// press, exactly as if the tab had never moved).

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { armConfirm, CONFIRM_MS } from "../../src/ui/confirm";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";

enableStrictLoadSemantics();

function danger(label: string): HTMLButtonElement {
  const button = document.createElement("button");
  button.textContent = label;
  return button;
}

describe("two independent armConfirm instances", () => {
  it("arming the second button never disarms, relabels or retimes the first", () => {
    vi.useFakeTimers();
    try {
      const buttonA = danger("Remove link");
      const buttonB = danger("Forget key");
      const runA = vi.fn();
      const runB = vi.fn();
      const confirmA = armConfirm(buttonA, runA);
      const confirmB = armConfirm(buttonB, runB);

      confirmA.handleClick(); // arm A
      vi.advanceTimersByTime(CONFIRM_MS - 500);
      confirmB.handleClick(); // arm B, well within A's own window

      expect(confirmA.isArmed()).toBe(true);
      expect(confirmB.isArmed()).toBe(true);
      expect(buttonA.textContent).toBe("Click again to confirm");
      expect(buttonB.textContent).toBe("Click again to confirm");

      // A's own timer was armed 500 ms before B's: it must lapse on ITS
      // schedule, unaffected by B ever having been touched.
      vi.advanceTimersByTime(500);
      expect(confirmA.isArmed()).toBe(false);
      expect(buttonA.textContent).toBe("Remove link");
      expect(confirmB.isArmed()).toBe(true); // B still has 500 ms left

      confirmB.handleClick();
      expect(runB).toHaveBeenCalledTimes(1);
      expect(runA).not.toHaveBeenCalled();
    } finally {
      vi.useRealTimers();
    }
  });
});

// A minimal boot, the same technique test/hunt/main.boot-edges.hunt.test.ts
// uses: two real data-confirm buttons on two different tabs (Reset all on
// Brand, Clean past the data on Workbook), wired by the real wireControls().
function pane(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

function patchOfficeForBoot(): void {
  const office = (
    globalThis as unknown as {
      Office: {
        onReady: (cb?: (info: { host: string }) => unknown) => unknown;
      };
    }
  ).Office;
  office.onReady = (callback) => {
    const info = { host: "Excel" };
    callback?.(info);
    return Promise.resolve(info);
  };
}

async function drain(rounds = 12): Promise<void> {
  for (let i = 0; i < rounds; i += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

async function bootExcel(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  pane();
  installFakeHost({ sheets: ["Model", "Data"] });
  patchOfficeForBoot();
  await import("../../src/main");
  await drain();
}

function clickTab(id: string): void {
  (document.getElementById(id) as HTMLButtonElement).click();
}

afterEach(() => {
  uninstallFakeHost();
  vi.restoreAllMocks();
});

describe("a tab switch while a real button sits armed", () => {
  it("clean-past-data stays armed across a switch to Brand and back to Workbook", async () => {
    await bootExcel();
    clickTab("tab-workbook");
    await drain();

    const button = document.querySelector(
      '[data-action="clean-past-data"]',
    ) as HTMLButtonElement;
    button.click(); // arm
    expect(button.classList.contains("armed")).toBe(true);

    clickTab("tab-brand");
    await drain();
    clickTab("tab-workbook");
    await drain();

    expect(button.classList.contains("armed")).toBe(true);
    expect(button.querySelector("strong")?.textContent).toBe(
      "Click again to confirm",
    );

    // The second press still runs the real action, exactly as if the tabs
    // had never moved.
    button.click();
    await drain();
    expect(button.classList.contains("armed")).toBe(false);
  });

  it("Reset all (Brand) and Clean past the data (Workbook) can be armed at once, independently", async () => {
    await bootExcel();
    clickTab("tab-brand");
    await drain();
    const resetAll = document.getElementById(
      "shortcuts-reset",
    ) as HTMLButtonElement;
    resetAll.click(); // arm, host-free action
    expect(resetAll.classList.contains("armed")).toBe(true);

    clickTab("tab-workbook");
    await drain();
    const cleanPastData = document.querySelector(
      '[data-action="clean-past-data"]',
    ) as HTMLButtonElement;
    cleanPastData.click(); // arm, on a different tab, while Reset all is still armed

    expect(cleanPastData.classList.contains("armed")).toBe(true);
    clickTab("tab-brand");
    await drain();
    // Still armed after the round trip, unaffected by the other button.
    expect(resetAll.classList.contains("armed")).toBe(true);
  });
});
