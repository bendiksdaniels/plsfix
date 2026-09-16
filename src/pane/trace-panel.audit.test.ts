// @vitest-environment jsdom
// The Audit section as the modeller drives it: the overlay toggle's state pill
// and note, and the Smart Track walk - chips, the jump each one makes, the back
// stack and its limit. The Excel adapter is mocked, so no Office host is needed;
// the markup under test is taskpane.html itself, which keeps ids from drifting.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { TraceArea, TraceResult } from "../excel";
import type * as ExcelShared from "../excel/shared";

// trace-panel reaches ./shared, whose top-level installTabs/createToast calls
// need #tab-bar and #toast: the markup goes in before the dynamic import.
vi.mock("../excel", async () => {
  const shared = await vi.importActual<typeof ExcelShared>("../excel/shared");
  return {
    parseAddress: shared.parseAddress,
    lastAuditNote: vi.fn(() => null),
    selectArea: vi.fn(async () => undefined),
    selectAreas: vi.fn(async () => undefined),
    toggleAuditOverlay: vi.fn(async () => true),
    traceActiveCell: vi.fn(async () => ({ origin: "Model!B2", areas: [] })),
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

// `hidden` is `boolean | "until-found"` in the DOM types, so it is read as the
// attribute the pane's [hidden] reset actually keys off.
function panelHidden(): boolean {
  return (
    document.getElementById("trace-panel")?.hasAttribute("hidden") ?? false
  );
}

function text(id: string): string {
  return document.getElementById(id)?.textContent ?? "";
}

// A chip click starts an action nobody awaits; the shared guard's own chain is
// a handful of microtasks long.
async function settle(): Promise<void> {
  for (let turn = 0; turn < 6; turn += 1) await Promise.resolve();
}

beforeEach(() => {
  vi.clearAllMocks();
});

describe("the audit overlay toggle", () => {
  it("says on, then off, in the pill beside the button", async () => {
    const { excel, panel } = await load();
    panel.renderAuditState();
    expect(text("audit-state")).toBe("Off");

    expect(await panel.toggleAudit()).toBe("Audit overlay on");
    expect(text("audit-state")).toBe("On");

    vi.mocked(excel.toggleAuditOverlay).mockResolvedValueOnce(false);
    expect(await panel.toggleAudit()).toBe("Audit overlay off");
    expect(text("audit-state")).toBe("Off");
  });

  it("carries the note the adapter left behind", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.toggleAuditOverlay).mockResolvedValueOnce(false);
    vi.mocked(excel.lastAuditNote).mockReturnValueOnce(
      "Audit overlay: this sheet is protected, nothing was changed",
    );

    expect(await panel.toggleAudit()).toBe(
      "Audit overlay off: Audit overlay: this sheet is protected, nothing was changed",
    );
    expect(text("audit-state")).toBe("Off");
  });
});

describe("the trace panel", () => {
  function traced(areas: TraceArea[], origin = "Model!B2"): TraceResult {
    return { origin, areas };
  }

  it("stays hidden until a trace has run", async () => {
    await load();
    expect(panelHidden()).toBe(true);
  });

  it("lists one chip per area and names the origin", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce(
      traced([area("Model", "A1:A3", 3), area("Data", "C1")]),
    );

    expect(await panel.startTrace("precedents", false)).toBe(
      "2 direct precedents",
    );
    expect(panelHidden()).toBe(false);
    expect(text("trace-origin")).toBe("Model!B2 · precedents");
    expect(chips().map((chip) => chip.textContent)).toEqual([
      "Model!A1:A3",
      "Data!C1",
    ]);
    expect(chips()[0]?.title).toBe("Select Model!A1:A3 (3 cells)");
    expect(
      (document.getElementById("trace-back") as HTMLButtonElement).disabled,
    ).toBe(true);
  });

  it("counts a whole-column precedent in the chip's own title", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce(
      traced([area("Model", "A:A", 1_048_576)]),
    );

    await panel.startTrace("precedents", false);
    expect(chips()[0]?.textContent).toBe("Model!A:A");
    expect(chips()[0]?.title).toBe("Select Model!A:A (1,048,576 cells)");
  });

  it("says one dependent in the singular", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce(
      traced([area("Model", "D9")]),
    );
    expect(await panel.startTrace("dependents", false)).toBe(
      "1 direct dependent",
    );
  });

  it("shows a line of its own when the cell reads from nothing", async () => {
    const { panel } = await load();
    expect(await panel.startTrace("precedents", false)).toBe(
      "No direct precedents",
    );
    expect(panelHidden()).toBe(false);
    expect(text("trace-chips")).toBe("No direct precedents.");
    expect(chips()).toHaveLength(0);
  });

  it("jumps to the first area when the keystroke asked for it", async () => {
    const { excel, panel } = await load();
    const first = area("Data", "C1");
    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce(traced([first]));

    await panel.startTrace("precedents", true);
    expect(excel.selectAreas).toHaveBeenCalledWith([first]);
  });

  it("jumps to every same-sheet area at once, cross-sheet ones left for the chips", async () => {
    const { excel, panel } = await load();
    const areas = [
      area("Model", "A1"),
      area("Model", "C3"),
      area("Data", "B2"),
    ];
    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce(traced(areas));

    await panel.startTrace("precedents", true);
    expect(excel.selectAreas).toHaveBeenCalledWith(areas);
    // Every area still lists as a chip regardless of which ones got selected.
    expect(chips().map((chip) => chip.textContent)).toEqual([
      "Model!A1",
      "Model!C3",
      "Data!B2",
    ]);
  });

  it("leaves the selection alone when the pane is open", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce(
      traced([area("Data", "C1")]),
    );

    await panel.startTrace("precedents", false);
    expect(excel.selectArea).not.toHaveBeenCalled();
    expect(excel.selectAreas).not.toHaveBeenCalled();
  });

  it("keeps no stale chips from the trace before", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce(
      traced([area("Model", "A1"), area("Model", "A2")]),
    );
    await panel.startTrace("precedents", false);

    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce(
      traced([area("Model", "Z9")]),
    );
    await panel.startTrace("dependents", false);
    expect(chips().map((chip) => chip.textContent)).toEqual(["Model!Z9"]);
  });
});

describe("walking the chain", () => {
  it("jumps where the chip points and traces on from there", async () => {
    const { excel, panel } = await load();
    const hop = area("Model", "A1:A3", 3);
    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce({
      origin: "Model!B2",
      areas: [hop],
    });
    await panel.startTrace("precedents", false);

    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce({
      origin: "Model!A1",
      areas: [area("Data", "C1")],
    });
    chips()[0]?.click();
    await settle();

    expect(excel.selectArea).toHaveBeenCalledWith(hop);
    expect(text("trace-origin")).toBe("Model!A1 · precedents");
    expect(
      (document.getElementById("trace-back") as HTMLButtonElement).disabled,
    ).toBe(false);
  });

  it("steps back to the cell it came from", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce({
      origin: "Model!B2",
      areas: [area("Model", "A1")],
    });
    await panel.startTrace("precedents", false);

    vi.mocked(excel.traceActiveCell).mockResolvedValueOnce({
      origin: "Model!A1",
      areas: [],
    });
    chips()[0]?.click();
    await settle();

    expect(await panel.traceBack()).toBe("Back at Model!B2");
    expect(excel.selectArea).toHaveBeenLastCalledWith({
      sheet: "Model",
      address: "B2",
    });
    expect(text("trace-origin")).toBe("Model!B2 · precedents");
    expect(
      (document.getElementById("trace-back") as HTMLButtonElement).disabled,
    ).toBe(true);
  });

  it("says so when there is nowhere to go back to", async () => {
    const { excel, panel } = await load();
    expect(await panel.traceBack()).toBe("Nothing to go back to");
    expect(excel.selectArea).not.toHaveBeenCalled();
  });

  it("a fresh trace empties the back stack", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.traceActiveCell).mockResolvedValue({
      origin: "Model!B2",
      areas: [area("Model", "A1")],
    });
    await panel.startTrace("precedents", false);
    chips()[0]?.click();
    await settle();

    await panel.startTrace("dependents", false);
    expect(await panel.traceBack()).toBe("Nothing to go back to");
  });

  it("keeps the back stack to its last twenty hops", async () => {
    const { excel, panel } = await load();
    vi.mocked(excel.traceActiveCell).mockImplementation(async () => ({
      origin: "Model!B2",
      areas: [area("Model", "A1")],
    }));
    await panel.startTrace("precedents", false);

    for (let hop = 0; hop < 25; hop += 1) {
      chips()[0]?.click();
      await settle();
    }

    let steps = 0;
    while ((await panel.traceBack()) !== "Nothing to go back to") steps += 1;
    expect(steps).toBe(20);
  });
});
