// @vitest-environment jsdom
// "Find a combination" as the modeller drives it: the two boxes, the result
// line in the house number style of the Brand language, and what the line says
// after a search the sheet refused. The Excel adapter is mocked, so no host is
// needed; the markup under test is taskpane.html itself.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { reconcileSelection } from "../excel";
import { DEFAULT_SETTINGS, setActiveSettings } from "../settings";
import { runReconciliation } from "./reconcile-panel";

vi.mock("../excel", () => ({ reconcileSelection: vi.fn() }));

const HINT =
  "Select up to 34 numeric cells. Matching cells become the selection.";

function paneRoot(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

function field(id: string): HTMLInputElement {
  return document.getElementById(id) as HTMLInputElement;
}

function resultLine(): string {
  return (
    document.getElementById("reconcile-result")?.textContent ?? ""
  ).trim();
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

beforeEach(() => {
  paneRoot();
  setActiveSettings(DEFAULT_SETTINGS);
  vi.mocked(reconcileSelection).mockReset();
});

describe("the reconcile panel", () => {
  it("passes both boxes to the adapter and reports the answer", async () => {
    field("reconcile-target").value = "1234.56";
    field("reconcile-tolerance").value = "0.5";
    vi.mocked(reconcileSelection).mockResolvedValue({
      addresses: ["B4", "B7", "B9"],
      count: 3,
      difference: 0,
      sum: 1234.56,
      values: [1000, 200, 34.56],
    });

    expect(await runReconciliation()).toBe("Found 3 matching cells");
    expect(reconcileSelection).toHaveBeenCalledWith(1234.56, 0.5);
    // Latvian is the pane's default: thousands grouped by a space.
    expect(resultLine()).toBe("3 cells selected · Sum 1 234.56 · Variance 0");
  });

  it("follows the Brand language", async () => {
    setActiveSettings({ ...DEFAULT_SETTINGS, language: "en" });
    field("reconcile-target").value = "98765";
    vi.mocked(reconcileSelection).mockResolvedValue({
      addresses: ["C2"],
      count: 5,
      difference: -12.34,
      sum: 98765,
      values: [98765],
    });

    await runReconciliation();
    expect(resultLine()).toBe(
      "5 cells selected · Sum 98,765 · Variance -12.34",
    );
  });

  it("leaves no stale answer standing after a refused search", async () => {
    field("reconcile-target").value = "100";
    vi.mocked(reconcileSelection).mockResolvedValue({
      addresses: ["A1", "A2"],
      count: 2,
      difference: 0,
      sum: 100,
      values: [60, 40],
    });
    await runReconciliation();
    expect(resultLine()).toContain("2 cells selected");

    // The next target reaches nothing: the cells the line names are not the
    // selection any more, so the line may not keep naming them.
    field("reconcile-target").value = "101";
    vi.mocked(reconcileSelection).mockRejectedValue(
      new Error("No combination reaches the target within the tolerance."),
    );
    expect(await rejects(() => runReconciliation())).toBe(
      "No combination reaches the target within the tolerance.",
    );
    expect(resultLine()).toBe(HINT);
  });

  it("refuses an empty box and a negative tolerance by name", async () => {
    field("reconcile-target").value = "";
    expect(await rejects(() => runReconciliation())).toBe(
      "Target must be a number.",
    );

    field("reconcile-target").value = "10";
    field("reconcile-tolerance").value = "";
    expect(await rejects(() => runReconciliation())).toBe(
      "Tolerance must be a number.",
    );

    field("reconcile-tolerance").value = "-1";
    expect(await rejects(() => runReconciliation())).toBe(
      "Tolerance must be zero or greater.",
    );
    expect(reconcileSelection).not.toHaveBeenCalled();
    // Nothing ran, so the standing hint is untouched.
    expect(resultLine()).toBe(HINT);
  });

  it("clears the old answer even when a box is not a number", async () => {
    field("reconcile-target").value = "100";
    vi.mocked(reconcileSelection).mockResolvedValue({
      addresses: ["A1", "A2"],
      count: 2,
      difference: 0,
      sum: 100,
      values: [60, 40],
    });
    await runReconciliation();
    expect(resultLine()).toContain("2 cells selected");

    // The refusal comes before any search: the line still may not keep naming
    // cells that answered an earlier target.
    field("reconcile-target").value = "hundred";
    expect(await rejects(() => runReconciliation())).toBe(
      "Target must be a number.",
    );
    expect(resultLine()).toBe(HINT);
  });
});
