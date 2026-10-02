// @vitest-environment jsdom
// The Workbook tab's two review panels: Prepare for sharing and the model
// check - what each row says, what the Copy button hands over and where a
// finding with no cell sends the reviewer instead. Super Find and the style
// scrubber are next door in workbook-panels.audit.test.ts.
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

import { jumpToHit, prepareForSharing, runModelCheck } from "../excel";
import type { Finding, ModelCheckReport } from "../model-check";
import type { ShareIssue } from "../share";

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
    share: await import("./share-panel"),
    check: await import("./model-check-panel"),
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

function finding(over: Partial<Finding> = {}): Finding {
  return {
    kind: "hardcodeInFormula",
    sheet: "Model",
    ref: "B4",
    count: 1,
    note: "=B3*1.1",
    ...over,
  };
}

function report(findings: Finding[]): ModelCheckReport {
  return {
    findings,
    scanned: { sheets: 2, cells: 40 },
    skipped: [],
    truncated: false,
  };
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

describe("share panel", () => {
  function issue(over: Partial<ShareIssue> = {}): ShareIssue {
    return { kind: "hiddenSheet", label: "Scratch", ...over };
  }

  it("shows the standing hint before the first run", async () => {
    const { share } = await load();
    share.renderShare(null);

    expect(text("share-hint")).toContain("Nothing is deleted");
    expect(document.getElementById("share-report")?.hidden).toBe(true);
  });

  it("badges every kind and leads with what the run did", async () => {
    vi.mocked(prepareForSharing).mockResolvedValue({
      report: [issue(), issue({ kind: "linkTokens", label: "Link registry" })],
      touchedSheets: 3,
      scannedSheets: 3,
      sheetCap: 200_000,
    });
    const { share } = await load();

    const line = await share.prepareShare();

    expect(line).toBe(
      "3 sheets reset to A1; 1 hidden sheet left; link tokens travel with the file",
    );
    expect(
      children("share-report").map(
        (row) => row.querySelector(".sheet-badge")?.textContent,
      ),
    ).toEqual(["Hidden", "Links"]);
    expect(text("share-hint")).toContain("Zoom cannot be reset");
  });

  it("names how many sheets it read, how many it skipped and over what cap", async () => {
    vi.mocked(prepareForSharing).mockResolvedValue({
      report: [issue({ kind: "skippedSheet", label: "Data" })],
      touchedSheets: 7,
      scannedSheets: 7,
      sheetCap: 200_000,
    });
    const { share } = await load();

    await share.prepareShare();

    expect(text("share-hint")).toContain(
      "Scanned 7 sheets, 1 skipped over 200,000 cells: Data.",
    );
  });

  it("cuts a long report off and says how much is left", async () => {
    vi.mocked(prepareForSharing).mockResolvedValue({
      report: Array.from({ length: 23 }, () => issue()),
      touchedSheets: 1,
      scannedSheets: 1,
      sheetCap: 200_000,
    });
    const { share } = await load();

    await share.prepareShare();

    const listed = children("share-report");
    expect(listed).toHaveLength(21);
    expect(listed[20]?.textContent).toBe("…and 3 more.");
  });
});

describe("model check panel", () => {
  it("keeps Copy out of reach until there is a report", async () => {
    const { check } = await load();
    check.renderModelCheck();

    const copy = document.getElementById(
      "copy-model-check",
    ) as HTMLButtonElement;
    expect(copy.disabled).toBe(true);
    expect(await check.copyReport()).toBe("Run the model check first.");

    vi.mocked(runModelCheck).mockResolvedValue(report([finding()]));
    expect(await check.runCheck()).toBe("Model check: 1 finding on 2 sheets");
    expect(copy.disabled).toBe(false);

    // jsdom does not define execCommand at all (so the real fallback throws,
    // caught, as false); assigning it stands in for a host where the copy
    // genuinely lands.
    Object.assign(document, { execCommand: vi.fn(() => true) });
    expect(await check.copyReport()).toBe("Copied 1 lines");
    Reflect.deleteProperty(document, "execCommand");
  });

  it("reports the copy failed when neither clipboard path works", async () => {
    const { check } = await load();
    vi.mocked(runModelCheck).mockResolvedValue(report([finding()]));
    await check.runCheck();

    Object.defineProperty(navigator, "clipboard", {
      value: undefined,
      configurable: true,
    });
    Object.assign(document, { execCommand: vi.fn(() => false) });
    await expect(check.copyReport()).rejects.toThrow(
      "Copy failed: select the text and copy it by hand.",
    );
    Reflect.deleteProperty(document, "execCommand");
  });

  it("names the location, the kind and the note on every row", async () => {
    vi.mocked(runModelCheck).mockResolvedValue(
      report([
        finding(),
        finding({ kind: "brokenName", sheet: null, ref: null, note: "Costs" }),
      ]),
    );
    const { check } = await load();

    await check.runCheck();

    const listed = children("model-check-list");
    expect(
      listed.map((row) => row.querySelector("strong")?.textContent),
    ).toEqual(["Model!B4", "Workbook"]);
    expect(
      listed.map((row) => row.querySelector(".sheet-badge")?.textContent),
    ).toEqual(["Hardcode in formula", "Broken name"]);
  });

  // A finding the whole file carries has no cell: the row says where the fix
  // lives instead of failing on a jump nobody can make.
  it("points at the section instead of jumping for a workbook-wide finding", async () => {
    vi.mocked(runModelCheck).mockResolvedValue(
      report([
        finding({ kind: "unusedStyle", sheet: null, ref: null, note: "Old" }),
        finding({ kind: "hiddenSheet", ref: null, note: "Hidden" }),
      ]),
    );
    const { check } = await load();
    await check.runCheck();
    const listed = children("model-check-list");

    (listed[1] as HTMLButtonElement).click();
    await settle();
    expect(text("toast")).toBe("Model has no cell to jump to.");

    (listed[0] as HTMLButtonElement).click();
    await settle();
    expect(text("toast")).toBe(
      "The Styles section above deletes unused styles.",
    );
    expect(jumpToHit).not.toHaveBeenCalled();
  });

  it("jumps to the cell a finding names and re-reads the explorer", async () => {
    vi.mocked(runModelCheck).mockResolvedValue(report([finding()]));
    const { check } = await load();
    await check.runCheck();

    (children("model-check-list")[0] as HTMLButtonElement).click();
    await settle();

    expect(jumpToHit).toHaveBeenCalledWith({
      kind: "cell",
      sheet: "Model",
      address: "B4",
      text: "=B3*1.1",
    });
    expect(text("toast")).toBe("Jumped to Model!B4");
  });

  it("cuts the list at forty and points the rest at the copied report", async () => {
    vi.mocked(runModelCheck).mockResolvedValue(
      report(Array.from({ length: 45 }, () => finding())),
    );
    const { check } = await load();

    await check.runCheck();

    const listed = children("model-check-list");
    expect(listed).toHaveLength(41);
    expect(listed[40]?.textContent).toBe("…and 5 more in the copied report.");
  });
});
