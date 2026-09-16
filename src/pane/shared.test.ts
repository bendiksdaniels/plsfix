// @vitest-environment jsdom
// src/pane/shared.ts has never run in a test (24.13%/0% branch coverage):
// the shared guard (busy toggle, decorate, describe, finally), the selection
// refresh and isExcelReady/setExcelReady. Driven over the real taskpane.html
// with ../excel mocked, so no Office host is needed. Each test reloads the
// module fresh (vi.resetModules) since actionButtons/toast/tabs are captured
// once, at import time, off whatever DOM is live then.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  copySourceLabel,
  inspectSelection,
  lastUndoSkipped,
  undoTarget,
} from "../excel";
import type * as SharedModule from "./shared";

vi.mock("../excel", () => ({
  copySourceLabel: vi.fn(() => null as string | null),
  inspectSelection: vi.fn(async () => ({
    address: "A1",
    cells: 1,
    formulas: 0,
    errors: 0,
    blanks: 0,
  })),
  lastUndoSkipped: vi.fn(() => false),
  undoTarget: vi.fn(() => null as string | null),
}));

async function load(): Promise<typeof SharedModule> {
  vi.resetModules();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  return import("./shared");
}

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

function toastClass(): string {
  return document.getElementById("toast")?.className ?? "";
}

beforeEach(() => {
  vi.mocked(lastUndoSkipped).mockReturnValue(false);
  vi.mocked(copySourceLabel).mockReturnValue(null);
  vi.mocked(undoTarget).mockReturnValue(null);
  vi.mocked(inspectSelection).mockResolvedValue({
    address: "A1",
    cells: 1,
    formulas: 0,
    errors: 0,
    blanks: 0,
  });
});

describe("refreshSelection", () => {
  it("renders the selection summary into the metric elements", async () => {
    vi.mocked(inspectSelection).mockResolvedValueOnce({
      address: "B2:C4",
      cells: 6,
      formulas: 2,
      errors: 1,
      blanks: 0,
    });
    const { refreshSelection } = await load();
    await refreshSelection();
    expect(document.getElementById("selection-address")?.textContent).toBe(
      "B2:C4",
    );
    expect(document.getElementById("metric-cells")?.textContent).toBe("6");
    expect(document.getElementById("metric-formulas")?.textContent).toBe("2");
    expect(document.getElementById("metric-errors")?.textContent).toBe("1");
    expect(document.getElementById("metric-blanks")?.textContent).toBe("0");
  });

  it("shows an em dash for a metric too large to read (-1)", async () => {
    vi.mocked(inspectSelection).mockResolvedValueOnce({
      address: "A:A",
      cells: -1,
      formulas: -1,
      errors: 0,
      blanks: 0,
    });
    const { refreshSelection } = await load();
    await refreshSelection();
    expect(document.getElementById("metric-cells")?.textContent).toBe("—");
    expect(document.getElementById("metric-formulas")?.textContent).toBe("—");
  });

  it("toasts the error instead of throwing when the read fails", async () => {
    vi.mocked(inspectSelection).mockRejectedValueOnce(new Error("boom"));
    const { refreshSelection } = await load();
    await expect(refreshSelection()).resolves.toBeUndefined();
    expect(toastText()).toBe("boom");
    expect(toastClass()).toContain("error");
  });

  // errorMessage()'s own fallback: a rejection that is not an Error instance
  // at all (a thrown string, say) still reads as a plain sentence, not
  // "undefined" or a raw dump of whatever was thrown.
  it("falls back to a plain sentence when the rejection is not an Error", async () => {
    vi.mocked(inspectSelection).mockRejectedValueOnce("not an Error object");
    const { refreshSelection } = await load();
    await expect(refreshSelection()).resolves.toBeUndefined();
    expect(toastText()).toBe("Excel could not complete that action.");
  });
});

describe("renderActionState", () => {
  it("shows the undo target and the marked copy source", async () => {
    vi.mocked(undoTarget).mockReturnValue("Model!A1:B2");
    vi.mocked(copySourceLabel).mockReturnValue("Model!C3");
    const { renderActionState } = await load();
    renderActionState();
    expect(document.getElementById("undo-target")?.textContent).toBe(
      "Model!A1:B2",
    );
    expect(document.getElementById("paste-source")?.textContent).toBe(
      "Copy source: Model!C3",
    );
  });

  it("falls back to placeholder text with nothing to report", async () => {
    const { renderActionState } = await load();
    renderActionState();
    expect(document.getElementById("undo-target")?.textContent).toBe(
      "Nothing to undo yet",
    );
    expect(document.getElementById("paste-source")?.textContent).toBe(
      "Mark a source, then paste it into any selection.",
    );
  });
});

describe("guard", () => {
  it("disables every [data-action] button for the run, re-enables after", async () => {
    const { guard } = await load();
    const buttons = Array.from(
      document.querySelectorAll<HTMLButtonElement>("[data-action]"),
    );
    expect(buttons.length).toBeGreaterThan(0);
    let sawBusy = false;
    await guard(async () => {
      sawBusy = buttons.every((button) => button.disabled);
      return "done";
    });
    expect(sawBusy).toBe(true);
    expect(buttons.every((button) => !button.disabled)).toBe(true);
  });

  it("shows the run's own message on success", async () => {
    const { guard } = await load();
    await guard(async () => "Saved");
    expect(toastText()).toBe("Saved");
    expect(toastClass()).not.toContain("error");
  });

  it("describes and toasts a failure instead of throwing out of the guard", async () => {
    const { guard } = await load();
    await expect(
      guard(async () => {
        throw new Error("Excel refused");
      }, "undo"),
    ).resolves.toBeUndefined();
    expect(toastText()).toBe("Undo: Excel refused");
    expect(toastClass()).toContain("error");
  });

  it("decorates a genuinely skipped undo on the success path", async () => {
    vi.mocked(lastUndoSkipped).mockReturnValueOnce(true);
    const { guard } = await load();
    await guard(async () => "Painted");
    expect(toastText()).toBe("Painted (too large for undo)");
  });

  // The bug: decorate (which drains lastUndoSkipped) only runs on success, so
  // a failed action that had skipped undo left the flag armed for whatever
  // succeeded next. finally must drain it every time, not just on success.
  it("does not leak a skipped-undo flag from a failed action into the next success", async () => {
    let armed = true;
    vi.mocked(lastUndoSkipped).mockImplementation(() => {
      const was = armed;
      armed = false;
      return was;
    });
    const { guard } = await load();

    await guard(async () => {
      throw new Error("cap exceeded");
    }, "cycle-fill");
    expect(toastText()).toBe("Cycle fill: cap exceeded");

    await guard(async () => "Selection updated");
    expect(toastText()).toBe("Selection updated");
  });
});

describe("isExcelReady / setExcelReady", () => {
  it("starts false and reflects whatever boot() sets", async () => {
    const { isExcelReady, setExcelReady } = await load();
    expect(isExcelReady()).toBe(false);
    setExcelReady(true);
    expect(isExcelReady()).toBe(true);
  });
});
