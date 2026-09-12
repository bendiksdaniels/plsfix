// @vitest-environment jsdom
// "Precedents of selection" as the modeller sees it: the trace panel's list
// grouped by source cell, and the one-line toast counting precedents, cells
// and the cells that hold no formula. The Excel adapter is mocked, so no
// Office host is needed; the markup is taskpane.html itself.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { MultiTraceResult, TraceArea } from "../excel";
import type * as ExcelShared from "../excel/shared";

vi.mock("../excel", async () => {
  const shared = await vi.importActual<typeof ExcelShared>("../excel/shared");
  return {
    parseAddress: shared.parseAddress,
    lastAuditNote: vi.fn(() => null),
    selectArea: vi.fn(async () => undefined),
    toggleAuditOverlay: vi.fn(async () => true),
    traceActiveCell: vi.fn(async () => ({ origin: "Model!B2", areas: [] })),
    tracePrecedentsOfSelection: vi.fn(),
    // Read by ./shared after every guarded action.
    copySourceLabel: vi.fn(() => null),
    lastUndoSkipped: vi.fn(() => false),
    undoTarget: vi.fn(() => null),
    inspectSelection: vi.fn(async () => ({
      address: "Model!B2",
      cells: 1,
      formulas: 0,
      errors: 0,
      blanks: 0,
    })),
  };
});

function paneRoot(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

// No explicit return type: annotating it would need an `import()` type, which
// the lint config forbids in favour of a top-level type import.
async function load() {
  vi.resetModules();
  paneRoot();
  const excel = await import("../excel");
  const panel = await import("./trace-panel");
  return { excel, panel };
}

function area(sheet: string, address: string, cellCount = 1): TraceArea {
  return { sheet, address, cellCount };
}

function chips(): HTMLButtonElement[] {
  return Array.from(
    document.querySelectorAll<HTMLButtonElement>("#trace-chips .chip"),
  );
}

function labels(): string[] {
  return Array.from(
    document.querySelectorAll<HTMLElement>("#trace-chips .hint"),
  ).map((line) => line.textContent ?? "");
}

function text(id: string): string {
  return document.getElementById(id)?.textContent ?? "";
}

function result(over: Partial<MultiTraceResult> = {}): MultiTraceResult {
  return {
    origin: "Model!B2:D2",
    groups: [],
    areas: [],
    formulaCells: 0,
    skipped: 0,
    ...over,
  };
}

beforeEach(() => {
  vi.clearAllMocks();
});

describe("the grouped precedents list", () => {
  it("shows one label per source cell with its own chips under it", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.tracePrecedentsOfSelection).mockResolvedValueOnce(
      result({
        groups: [
          { cell: "B2", areas: [area("Model", "A1:A3", 3)] },
          { cell: "C2", areas: [] },
          { cell: "D2", areas: [area("Data", "C1"), area("Model", "F9")] },
        ],
        areas: [
          area("Model", "A1:A3", 3),
          area("Data", "C1"),
          area("Model", "F9"),
        ],
        formulaCells: 3,
      }),
    );

    expect(await panel.startPrecedentsOfSelection()).toBe(
      "3 precedents of 3 cells",
    );
    expect(text("trace-origin")).toBe("Model!B2:D2 · precedents");
    expect(labels()).toEqual(["B2", "C2: no precedents", "D2"]);
    expect(chips().map((chip) => chip.textContent)).toEqual([
      "Model!A1:A3",
      "Data!C1",
      "Model!F9",
    ]);
    expect(
      (document.getElementById("trace-back") as HTMLButtonElement).disabled,
    ).toBe(true);
  });

  it("counts one precedent of one cell in the singular", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.tracePrecedentsOfSelection).mockResolvedValueOnce(
      result({
        groups: [{ cell: "B2", areas: [area("Model", "A1")] }],
        areas: [area("Model", "A1")],
        formulaCells: 1,
      }),
    );

    expect(await panel.startPrecedentsOfSelection()).toBe(
      "1 precedent of 1 cell",
    );
  });

  it("names the cells that hold no formula in the same line", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.tracePrecedentsOfSelection).mockResolvedValueOnce(
      result({
        groups: [{ cell: "B2", areas: [area("Model", "A1")] }],
        areas: [area("Model", "A1")],
        formulaCells: 6,
        skipped: 2,
      }),
    );

    expect(await panel.startPrecedentsOfSelection()).toBe(
      "1 precedent of 6 cells; 2 cells have no formula",
    );
  });

  it("says so when every formula in the selection reads from nothing", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.tracePrecedentsOfSelection).mockResolvedValueOnce(
      result({
        groups: [{ cell: "B2", areas: [] }],
        formulaCells: 1,
        skipped: 1,
      }),
    );

    expect(await panel.startPrecedentsOfSelection()).toBe(
      "No precedents of 1 cell; 1 cell has no formula",
    );
    expect(chips()).toHaveLength(0);
  });

  it("says so when the selection holds no formula at all", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.tracePrecedentsOfSelection).mockResolvedValueOnce(
      result({ skipped: 4 }),
    );

    expect(await panel.startPrecedentsOfSelection()).toBe(
      "No formulas in the selection.",
    );
    expect(labels()).toEqual(["No formulas in the selection."]);
  });

  it("keeps walking from a chip, one cell at a time", async () => {
    const { excel, panel } = await load();
    const hop = area("Model", "A1:A3", 3);
    vi.mocked(excel.tracePrecedentsOfSelection).mockResolvedValueOnce(
      result({
        groups: [{ cell: "B2", areas: [hop] }],
        areas: [hop],
        formulaCells: 1,
      }),
    );
    await panel.startPrecedentsOfSelection();

    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce({
      origin: "Model!A1",
      areas: [area("Data", "C1")],
    });
    chips()[0]?.click();
    for (let turn = 0; turn < 6; turn += 1) await Promise.resolve();

    expect(excel.selectArea).toHaveBeenCalledWith(hop);
    expect(text("trace-origin")).toBe("Model!A1 · precedents");
    // Back on a one-cell walk, so the grouped labels are gone.
    expect(labels()).toEqual([]);
  });

  it("empties the back stack, the way a fresh single-cell trace does", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.traceActiveCell).mockResolvedValue({
      origin: "Model!B2",
      areas: [area("Model", "A1")],
    });
    await panel.startTrace("precedents", false);
    chips()[0]?.click();
    for (let turn = 0; turn < 6; turn += 1) await Promise.resolve();

    vi.mocked(excel.tracePrecedentsOfSelection).mockResolvedValueOnce(
      result({ skipped: 1 }),
    );
    await panel.startPrecedentsOfSelection();
    expect(await panel.traceBack()).toBe("Nothing to go back to");
  });
});
