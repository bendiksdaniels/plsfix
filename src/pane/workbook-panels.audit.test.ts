// @vitest-environment jsdom
// Two of the Workbook tab's list panels: Super Find and the style scrubber.
// Driven through the real pane markup with the Excel adapter mocked, so no
// Office host is needed. Prepare for sharing and the model check are next
// door in workbook-review.audit.test.ts.
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
  listSheets: vi.fn(async () => [] as unknown[]),
  activateSheet: vi.fn(async () => undefined),
  deleteBrokenNames: vi.fn(async () => 0),
  insertToc: vi.fn(async () => undefined),
  listBrokenNames: vi.fn(async () => [] as string[]),
  setSheetVisibility: vi.fn(async () => undefined),
  findInWorkbook: vi.fn(),
  jumpToHit: vi.fn(async () => undefined),
  listUnusedStyles: vi.fn(),
  deleteUnusedStyles: vi.fn(async () => 0),
  prepareForSharing: vi.fn(),
  runModelCheck: vi.fn(),
}));

import {
  deleteUnusedStyles,
  findInWorkbook,
  jumpToHit,
  listUnusedStyles,
} from "../excel";
import type { FindHit, FindResult, StyleScan } from "../excel";

function paneRoot(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

async function load() {
  vi.resetModules();
  paneRoot();
  await import("./shared");
  return {
    find: await import("./find-panel"),
    styles: await import("./styles-panel"),
  };
}

function text(id: string): string {
  return document.getElementById(id)?.textContent ?? "";
}

function children(id: string): HTMLElement[] {
  return Array.from(
    (document.getElementById(id) as HTMLElement).children,
  ) as HTMLElement[];
}

function hit(over: Partial<FindHit> = {}): FindHit {
  return {
    kind: "cell",
    sheet: "Model",
    address: "B4",
    text: "Total",
    ...over,
  };
}

function found(over: Partial<FindResult> = {}): FindResult {
  return { hits: [], skippedSheets: [], commentsSkipped: false, ...over };
}

function scan(over: Partial<StyleScan> = {}): StyleScan {
  return { unused: [], total: 4, skippedSheets: [], ...over };
}

async function settle(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

beforeEach(() => {
  vi.clearAllMocks();
  // showAsTaskpane is core Office chrome the three ribbon entry points call.
  vi.stubGlobal("Office", { addin: { showAsTaskpane: async () => undefined } });
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.useRealTimers();
});

describe("super find panel", () => {
  it("treats an empty box as a note, not a failure", async () => {
    const { find } = await load();

    expect(await find.runFind()).toBe("Type something to find first.");
    expect(findInWorkbook).not.toHaveBeenCalled();
    expect(document.getElementById("find-results")?.hidden).toBe(true);
  });

  it("passes the three tick boxes straight through", async () => {
    vi.mocked(findInWorkbook).mockResolvedValue(found());
    const { find } = await load();
    (document.getElementById("find-query") as HTMLInputElement).value =
      "  margin  ";
    (document.getElementById("find-case") as HTMLInputElement).checked = true;
    (document.getElementById("find-comments") as HTMLInputElement).checked =
      false;

    await find.runFind();

    expect(findInWorkbook).toHaveBeenCalledWith("margin", {
      matchCase: true,
      inFormulas: true,
      inComments: false,
    });
  });

  it("labels each kind of hit and jumps from the row", async () => {
    vi.mocked(findInWorkbook).mockResolvedValue(
      found({
        hits: [
          hit(),
          hit({ kind: "name", sheet: "", address: "TaxRate", text: "=0.21" }),
          hit({ kind: "comment", sheet: "Data", address: "C3", text: "Check" }),
        ],
      }),
    );
    const { find } = await load();
    (document.getElementById("find-query") as HTMLInputElement).value = "x";

    expect(await find.runFind()).toBe("3 hits.");
    const listed = children("find-results");
    expect(
      listed.map((row) => row.querySelector("strong")?.textContent),
    ).toEqual(["Model!B4", "Name · TaxRate", "Comment · Data!C3"]);

    (listed[2] as HTMLButtonElement).click();
    await settle();
    expect(jumpToHit).toHaveBeenCalledWith(
      hit({ kind: "comment", sheet: "Data", address: "C3", text: "Check" }),
    );
  });

  // An old host has no comment collection at all, so "no matches" would read
  // as "nothing was written there" rather than "nobody looked".
  it("says what it could not read and what the host cannot search", async () => {
    vi.mocked(findInWorkbook).mockResolvedValue(
      found({ skippedSheets: ["Data", "Notes"], commentsSkipped: true }),
    );
    const { find } = await load();
    (document.getElementById("find-query") as HTMLInputElement).value = "x";

    await find.runFind();

    expect(text("find-hint")).toBe(
      "No matches. Too large to search: Data, Notes. Comments need Excel 365.",
    );
  });

  it("says the list is the first page once the cap is reached", async () => {
    const { FIND_HIT_CAP } = await import("../find");
    vi.mocked(findInWorkbook).mockResolvedValue(
      found({ hits: Array.from({ length: FIND_HIT_CAP }, () => hit()) }),
    );
    const { find } = await load();
    (document.getElementById("find-query") as HTMLInputElement).value = "x";

    expect(await find.runFind()).toBe(
      `200 hits (first ${String(FIND_HIT_CAP)}).`,
    );
  });

  // A formula runs longer than the row; the whole of it stays in the tooltip.
  it("clips a long line into the row and keeps it whole in the title", async () => {
    const long = `=SUM(${"A1,".repeat(60)}A2)`;
    vi.mocked(findInWorkbook).mockResolvedValue(
      found({ hits: [hit({ text: long })] }),
    );
    const { find } = await load();
    (document.getElementById("find-query") as HTMLInputElement).value = "SUM";

    await find.runFind();

    const row = children("find-results")[0];
    expect(row?.querySelector("small")?.textContent).toBe(
      `${long.slice(0, 90)}…`,
    );
    expect(row?.getAttribute("title")).toContain(long);
  });

  it("opens the pane on the Workbook tab with the caret in the box", async () => {
    const { find } = await load();
    const box = document.getElementById("find-query") as HTMLInputElement;
    box.value = "margin";

    expect(await find.focusFind()).toBe("Find ready");

    expect(document.getElementById("view-workbook")?.hidden).toBe(false);
    expect(document.activeElement).toBe(box);
    expect(box.selectionStart).toBe(0);
    expect(box.selectionEnd).toBe("margin".length);
  });
});

describe("style scrubber panel", () => {
  it("counts the unused styles out of the whole table", async () => {
    vi.mocked(listUnusedStyles).mockResolvedValue(
      scan({ unused: ["Old header"], total: 5 }),
    );
    const { styles } = await load();

    expect(await styles.scanStyles()).toBe(
      "1 unused style of 5 in the workbook.",
    );
    expect(children("styles-list")).toHaveLength(1);
    expect(
      (document.getElementById("styles-delete") as HTMLButtonElement).hidden,
    ).toBe(false);
  });

  // A style worn only on a sheet nobody could read would look unused, so the
  // count says it is incomplete and the delete stays out of reach.
  it("keeps the delete out of reach while a sheet went unread", async () => {
    vi.mocked(listUnusedStyles).mockResolvedValue(
      scan({ unused: ["Old header"], skippedSheets: ["Data"] }),
    );
    const { styles } = await load();

    await styles.scanStyles();

    const button = document.getElementById(
      "styles-delete",
    ) as HTMLButtonElement;
    expect(button.hidden).toBe(false);
    expect(button.disabled).toBe(true);
    expect(text("styles-result")).toContain(
      "Some sheets were too large to scan: Data.",
    );
  });

  it("arms, lapses, and rescans after a delete", async () => {
    vi.useFakeTimers();
    vi.mocked(listUnusedStyles).mockResolvedValue(scan({ unused: ["Old"] }));
    vi.mocked(deleteUnusedStyles).mockResolvedValue(1);
    const { styles } = await load();
    const shared = await import("./shared");
    await styles.scanStyles();

    styles.armStyles();
    expect(styles.isStylesArmed()).toBe(true);
    vi.advanceTimersByTime(shared.DELETE_CONFIRM_MS);
    expect(styles.isStylesArmed()).toBe(false);
    expect(deleteUnusedStyles).not.toHaveBeenCalled();

    vi.mocked(listUnusedStyles).mockResolvedValue(scan({ total: 3 }));
    expect(await styles.deleteStyles()).toBe("Deleted 1 unused style");
    expect(deleteUnusedStyles).toHaveBeenCalledWith(["Old"]);
    expect(text("styles-result")).toBe(
      "No unused custom styles of 3 in the workbook.",
    );
  });

  it("opens the pane on the Workbook tab and scans from the ribbon", async () => {
    vi.mocked(listUnusedStyles).mockResolvedValue(scan());
    const { styles } = await load();

    await styles.focusStyles();

    expect(listUnusedStyles).toHaveBeenCalledTimes(1);
    expect(document.getElementById("view-workbook")?.hidden).toBe(false);
  });

  // The delete acts on a list, never on a guess: before the first scan there
  // is no list and no button.
  it("shows nothing to act on before the first scan", async () => {
    const { styles } = await load();

    styles.renderStyles();

    expect(text("styles-result")).toBe("Not scanned yet.");
    expect(document.getElementById("styles-list")?.hidden).toBe(true);
    expect(
      (document.getElementById("styles-delete") as HTMLButtonElement).hidden,
    ).toBe(true);
  });
});
