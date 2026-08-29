import { describe, expect, it } from "vitest";
import {
  AUTOCOLOR_LABEL,
  isExternalFormula,
  type ShareIssue,
  type ShareScan,
  shareReport,
  summarizeShare,
} from "./share";

const EMPTY: ShareScan = {
  sheets: [],
  externalLinks: [],
  brokenNames: [],
  skippedSheets: [],
  autocolorOnEdit: false,
};

function scan(part: Partial<ShareScan>): ShareScan {
  return { ...EMPTY, ...part };
}

function issues(kind: ShareIssue["kind"], count: number): ShareIssue[] {
  return Array.from({ length: count }, (_unused, index) => ({
    kind,
    label: `${kind} ${index + 1}`,
  }));
}

describe("isExternalFormula", () => {
  it("catches a reference to another workbook", () => {
    expect(isExternalFormula("=[Budget.xlsx]Model!$B$4")).toBe(true);
    expect(isExternalFormula("='C:\\models\\[Budget.xlsx]Model'!$B$4")).toBe(
      true,
    );
  });

  it("leaves a plain, cross-sheet or hardcoded formula alone", () => {
    expect(isExternalFormula("=Model!$B$4*2")).toBe(false);
    expect(isExternalFormula("=SUM(B4:B9)")).toBe(false);
    expect(isExternalFormula("=B4*1.2")).toBe(false);
  });

  it("is not fooled by brackets inside text or by a value", () => {
    expect(isExternalFormula('="[not a link]"')).toBe(false);
    expect(isExternalFormula("[Budget.xlsx]")).toBe(false);
    expect(isExternalFormula(42)).toBe(false);
    expect(isExternalFormula(null)).toBe(false);
  });

  it("needs both brackets, so a half-typed formula is not a link", () => {
    expect(isExternalFormula("=[Budget")).toBe(false);
  });
});

describe("shareReport", () => {
  it("reports hidden sheets and names the very hidden ones", () => {
    expect(
      shareReport(
        scan({
          sheets: [
            { name: "Model", visibility: "Visible" },
            { name: "Scratch", visibility: "Hidden" },
            { name: "Archive", visibility: "VeryHidden" },
          ],
        }),
      ),
    ).toEqual([
      { kind: "hiddenSheet", label: "Scratch" },
      { kind: "hiddenSheet", label: "Archive (very hidden)" },
    ]);
  });

  it("reports an external link by cell and formula", () => {
    expect(
      shareReport(
        scan({
          externalLinks: [
            {
              sheet: "Model",
              address: "B4",
              formula: "=[Budget.xlsx]Model!$B$4",
            },
          ],
        }),
      ),
    ).toEqual([
      {
        kind: "externalLink",
        label: "Model!B4: =[Budget.xlsx]Model!$B$4",
      },
    ]);
  });

  it("reports broken names, skipped sheets and autocolor on edit", () => {
    expect(
      shareReport(
        scan({
          brokenNames: ["Costs"],
          skippedSheets: ["Data"],
          autocolorOnEdit: true,
        }),
      ),
    ).toEqual([
      { kind: "brokenName", label: "Costs" },
      { kind: "skippedSheet", label: "Data" },
      { kind: "autocolorOnEdit", label: AUTOCOLOR_LABEL },
    ]);
  });

  it("reports nothing for a workbook with nothing to report", () => {
    expect(
      shareReport(scan({ sheets: [{ name: "Model", visibility: "Visible" }] })),
    ).toEqual([]);
  });

  it("lists in workbook order: sheets, links, names, skipped, autocolor", () => {
    const report = shareReport({
      sheets: [{ name: "Scratch", visibility: "Hidden" }],
      externalLinks: [
        { sheet: "Model", address: "B4", formula: "=[a.xlsx]S!A1" },
      ],
      brokenNames: ["Costs"],
      skippedSheets: ["Data"],
      autocolorOnEdit: true,
    });

    expect(report.map((issue) => issue.kind)).toEqual([
      "hiddenSheet",
      "externalLink",
      "brokenName",
      "skippedSheet",
      "autocolorOnEdit",
    ]);
  });
});

describe("summarizeShare", () => {
  it("counts what is left in one line", () => {
    const report = [...issues("hiddenSheet", 2), ...issues("externalLink", 1)];

    expect(summarizeShare(report, 6)).toBe(
      "6 sheets reset to A1; 2 hidden sheets, 1 external link left",
    );
  });

  it("says so when there is nothing left to fix", () => {
    expect(summarizeShare([], 4)).toBe(
      "4 sheets reset to A1; nothing else to fix",
    );
  });

  it("keeps the singular for one sheet and one of each finding", () => {
    const report = [
      ...issues("hiddenSheet", 1),
      ...issues("brokenName", 1),
      ...issues("skippedSheet", 1),
    ];

    expect(summarizeShare(report, 1)).toBe(
      "1 sheet reset to A1; 1 hidden sheet, 1 broken name, 1 sheet too large to scan left",
    );
  });

  it("pluralizes every finding", () => {
    const report = [
      ...issues("externalLink", 3),
      ...issues("brokenName", 2),
      ...issues("skippedSheet", 2),
    ];

    expect(summarizeShare(report, 12)).toBe(
      "12 sheets reset to A1; 3 external links, 2 broken names, 2 sheets too large to scan left",
    );
  });

  it("gives autocolor on edit its own clause", () => {
    expect(summarizeShare(issues("autocolorOnEdit", 1), 3)).toBe(
      "3 sheets reset to A1; autocolor on edit is still on",
    );
    expect(
      summarizeShare(
        [...issues("hiddenSheet", 1), ...issues("autocolorOnEdit", 1)],
        3,
      ),
    ).toBe(
      "3 sheets reset to A1; 1 hidden sheet left; autocolor on edit is still on",
    );
  });
});
