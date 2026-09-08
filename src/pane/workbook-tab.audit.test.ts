// @vitest-environment jsdom
// The Workbook tab's sheet explorer and broken-name scrubber driven through
// the real pane markup: what each row offers, what a refused list says, and
// the two clicks the delete needs. The Excel adapter is mocked, so no Office
// host is needed.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../excel", () => ({
  copySourceLabel: vi.fn(() => null),
  inspectSelection: vi.fn(async () => ({
    address: "A1",
    cells: 1,
    formulas: 0,
    errors: 0,
    blanks: 0,
  })),
  lastUndoSkipped: vi.fn(() => false),
  undoTarget: vi.fn(() => null),
  activateSheet: vi.fn(async () => undefined),
  deleteBrokenNames: vi.fn(async () => 0),
  insertToc: vi.fn(async () => undefined),
  listBrokenNames: vi.fn(async () => [] as string[]),
  listSheets: vi.fn(async () => [] as unknown[]),
  setSheetVisibility: vi.fn(async () => undefined),
}));

import {
  activateSheet,
  deleteBrokenNames,
  insertToc,
  listBrokenNames,
  listSheets,
  setSheetVisibility,
} from "../excel";

function paneRoot(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

// ./shared reads the DOM at import time, so the markup goes in first and the
// module registry is reset between tests.
async function load() {
  vi.resetModules();
  paneRoot();
  const shared = await import("./shared");
  const tab = await import("./workbook-tab");
  return { shared, tab };
}

function sheet(name: string, visibility = "Visible", active = false) {
  return { name, visibility, active };
}

function rows(): HTMLElement[] {
  return Array.from(
    document.querySelectorAll<HTMLElement>("#sheet-list .sheet-row"),
  );
}

function toastText(): string {
  return document.getElementById("toast")?.textContent ?? "";
}

function deleteButton(): HTMLButtonElement {
  return document.getElementById("delete-names") as HTMLButtonElement;
}

async function settle(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(listSheets).mockResolvedValue([]);
  vi.mocked(listBrokenNames).mockResolvedValue([]);
});

afterEach(() => {
  vi.useRealTimers();
});

describe("sheet explorer", () => {
  it("asks for a connection instead of an empty list before Excel is there", async () => {
    const { tab } = await load();

    await tab.refreshSheets();

    expect(listSheets).not.toHaveBeenCalled();
    expect(document.getElementById("sheet-list")?.textContent).toBe(
      "Connect to Excel to list the sheets.",
    );
  });

  it("gives a visible sheet a button, a hidden one a label and an eye", async () => {
    vi.mocked(listSheets).mockResolvedValue([
      sheet("Model", "Visible", true),
      sheet("Scratch", "Hidden"),
      sheet("Archive", "VeryHidden"),
    ]);
    const { shared, tab } = await load();
    shared.setExcelReady(true);

    await tab.refreshSheets();

    const [model, scratch, archive] = rows();
    expect(model?.querySelector("button.sheet-name")).not.toBeNull();
    expect(model?.getAttribute("aria-current")).toBe("true");
    expect(scratch?.querySelector("button.sheet-name")).toBeNull();
    expect(scratch?.querySelector("span.sheet-name")?.textContent).toBe(
      "Scratch",
    );
    expect(scratch?.querySelector("button.eye")).not.toBeNull();
    // Very hidden is set outside Excel's UI: no name button and no eye.
    expect(archive?.classList.contains("locked")).toBe(true);
    expect(archive?.querySelector("button")).toBeNull();
    expect(archive?.querySelector(".sheet-badge")?.textContent).toBe(
      "Very hidden",
    );
  });

  it("goes to a sheet from its name and re-reads the list", async () => {
    vi.mocked(listSheets).mockResolvedValue([sheet("Model"), sheet("Data")]);
    const { shared, tab } = await load();
    shared.setExcelReady(true);
    await tab.refreshSheets();

    rows()[1]?.querySelector<HTMLButtonElement>("button.sheet-name")?.click();
    await settle();

    expect(activateSheet).toHaveBeenCalledWith("Data");
    expect(listSheets).toHaveBeenCalledTimes(2);
    expect(toastText()).toContain("Switched to Data");
  });

  it("hides a visible sheet and shows a hidden one from the same eye", async () => {
    vi.mocked(listSheets).mockResolvedValue([
      sheet("Model"),
      sheet("Scratch", "Hidden"),
    ]);
    const { shared, tab } = await load();
    shared.setExcelReady(true);
    await tab.refreshSheets();

    rows()[0]?.querySelector<HTMLButtonElement>("button.eye")?.click();
    await settle();
    expect(setSheetVisibility).toHaveBeenLastCalledWith("Model", false);

    rows()[1]?.querySelector<HTMLButtonElement>("button.eye")?.click();
    await settle();
    expect(setSheetVisibility).toHaveBeenLastCalledWith("Scratch", true);
  });

  it("says why the list is empty when the host refuses it", async () => {
    vi.mocked(listSheets).mockRejectedValue(
      new Error("Excel is busy with something else."),
    );
    const { shared, tab } = await load();
    shared.setExcelReady(true);

    await tab.refreshSheets();

    expect(rows()).toHaveLength(0);
    expect(toastText()).toContain("Excel is busy");
  });

  it("rebuilds the rows rather than stacking a second set on them", async () => {
    vi.mocked(listSheets).mockResolvedValue([sheet("Model")]);
    const { shared, tab } = await load();
    shared.setExcelReady(true);

    await tab.refreshSheets();
    await tab.refreshSheets();

    expect(rows()).toHaveLength(1);
  });
});

describe("broken-name scrubber", () => {
  it("counts what it found and offers the delete only then", async () => {
    const { tab } = await load();
    tab.renderNames(false);
    expect(document.getElementById("names-result")?.textContent).toBe(
      "Not scanned yet.",
    );
    expect(deleteButton().hidden).toBe(true);

    vi.mocked(listBrokenNames).mockResolvedValue(["Costs", "Margin"]);
    expect(await tab.scanNames()).toBe("2 broken names");
    expect(document.getElementById("names-result")?.textContent).toBe(
      "2 broken names: Costs, Margin",
    );
    expect(deleteButton().hidden).toBe(false);
    expect(deleteButton().textContent).toBe("Delete 2 broken names");
  });

  it("says so, and hides the delete, when there is nothing broken", async () => {
    const { tab } = await load();

    expect(await tab.scanNames()).toBe("No broken names");
    expect(document.getElementById("names-result")?.textContent).toBe(
      "No broken names.",
    );
    expect(deleteButton().hidden).toBe(true);
  });

  it("cuts a long list off at six with an ellipsis", async () => {
    vi.mocked(listBrokenNames).mockResolvedValue(
      Array.from({ length: 8 }, (_, at) => `Name${String(at)}`),
    );
    const { tab } = await load();

    await tab.scanNames();

    expect(document.getElementById("names-result")?.textContent).toBe(
      "8 broken names: Name0, Name1, Name2, Name3, Name4, Name5, …",
    );
  });

  // Deleting a name is outside both undo stacks, so the first click only arms
  // the button and the arming lapses on its own.
  it("arms, lapses and disarms without deleting anything", async () => {
    vi.useFakeTimers();
    vi.mocked(listBrokenNames).mockResolvedValue(["Costs"]);
    const { shared, tab } = await load();
    await tab.scanNames();

    expect(tab.isDeleteArmed()).toBe(false);
    tab.armDelete();
    expect(deleteButton().textContent).toBe("Click again to confirm");
    expect(deleteButton().classList.contains("armed")).toBe(true);

    vi.advanceTimersByTime(shared.DELETE_CONFIRM_MS);
    expect(tab.isDeleteArmed()).toBe(false);
    expect(deleteButton().textContent).toBe("Delete 1 broken name");
    expect(deleteBrokenNames).not.toHaveBeenCalled();
  });

  it("reports what the delete removed and empties the list", async () => {
    vi.mocked(listBrokenNames).mockResolvedValue(["Costs", "Margin"]);
    vi.mocked(deleteBrokenNames).mockResolvedValue(2);
    const { tab } = await load();
    await tab.scanNames();

    expect(await tab.deleteNames()).toBe("Deleted 2 broken names");
    expect(document.getElementById("names-result")?.textContent).toBe(
      "No broken names.",
    );
    expect(deleteButton().hidden).toBe(true);
  });
});

describe("contents sheet", () => {
  it("rewrites the sheet and re-reads the explorer behind it", async () => {
    vi.mocked(listSheets).mockResolvedValue([sheet("TOC"), sheet("Model")]);
    const { shared, tab } = await load();
    shared.setExcelReady(true);

    expect(await tab.insertTocSheet()).toBe("Contents sheet updated");

    expect(insertToc).toHaveBeenCalledTimes(1);
    expect(listSheets).toHaveBeenCalledTimes(1);
    expect(rows()).toHaveLength(2);
  });
});
