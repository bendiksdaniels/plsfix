// @vitest-environment jsdom
// The model check panel's skipped-sheets sentence. Super Find and prepare for
// sharing both name how many sheets were scanned, how many were skipped and
// over what cap ("Searched 8 sheets, 1 skipped over 200,000 cells: Notes.").
// The model check panel's own hint carried only a bare list of names, with no
// count and no cap - untested (every existing report() fixture in
// src/pane/workbook-review.audit.test.ts passes skipped: []), and named as
// still-old wording in CLAUDE.md's symptom index.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../../src/excel", () => ({
  MODEL_CHECK_CELL_CAP: 200_000,
  jumpToHit: vi.fn(async () => undefined),
  runModelCheck: vi.fn(),
}));

import { runModelCheck } from "../../src/excel";
import type { Finding, ModelCheckReport } from "../../src/model-check";

function paneRoot(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

async function load() {
  vi.resetModules();
  paneRoot();
  await import("../../src/pane/shared");
  return import("../../src/pane/model-check-panel");
}

function text(id: string): string {
  return document.getElementById(id)?.textContent ?? "";
}

function report(over: Partial<ModelCheckReport> = {}): ModelCheckReport {
  return {
    findings: [],
    scanned: { sheets: 2, cells: 40 },
    skipped: [],
    truncated: false,
    ...over,
  };
}

beforeEach(() => {
  vi.clearAllMocks();
});

describe("model check panel: skipped-sheets wording", () => {
  it("names how many sheets it read, how many it skipped and over what cap", async () => {
    vi.mocked(runModelCheck).mockResolvedValue(
      report({
        scanned: { sheets: 7, cells: 1_000 },
        skipped: ["Data"],
      }),
    );
    const check = await load();

    await check.runCheck();

    expect(text("model-check-hint")).toContain(
      "Scanned 7 sheets, 1 skipped over 200,000 cells: Data.",
    );
    expect(text("model-check-hint")).not.toContain("Too large to read");
  });

  it("names every skipped sheet and pluralises the count for more than one", async () => {
    vi.mocked(runModelCheck).mockResolvedValue(
      report({
        scanned: { sheets: 5, cells: 2_000 },
        skipped: ["Data", "Archive"],
      }),
    );
    const check = await load();

    await check.runCheck();

    expect(text("model-check-hint")).toContain(
      "Scanned 5 sheets, 2 skipped over 200,000 cells: Data, Archive.",
    );
  });

  it("adds no skipped-sheets sentence at all when nothing was skipped", async () => {
    vi.mocked(runModelCheck).mockResolvedValue(
      report({ scanned: { sheets: 3, cells: 40 } }),
    );
    const check = await load();

    await check.runCheck();

    const hint = text("model-check-hint");
    expect(hint).not.toContain("Scanned");
    expect(hint).not.toContain("skipped");
    expect(hint.startsWith("Model check: nothing to flag on 3 sheets.")).toBe(
      true,
    );
  });

  it("keeps the finding count and the skipped-sheets sentence together", async () => {
    const finding: Finding = {
      kind: "hardcodeInFormula",
      sheet: "Model",
      ref: "B4",
      count: 1,
      note: "=B3*1.1",
    };
    vi.mocked(runModelCheck).mockResolvedValue(
      report({
        findings: [finding],
        scanned: { sheets: 4, cells: 500 },
        skipped: ["Archive"],
      }),
    );
    const check = await load();

    const line = await check.runCheck();

    // The toast (summarize()'s own line) keeps its terser parenthetical; the
    // richer sentence lives only in the pane's own hint text.
    expect(line).toBe("Model check: 1 finding on 4 sheets (1 sheet skipped)");
    expect(text("model-check-hint")).toContain(
      "Scanned 4 sheets, 1 skipped over 200,000 cells: Archive.",
    );
  });
});
