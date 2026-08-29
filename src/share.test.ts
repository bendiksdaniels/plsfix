import { describe, expect, it } from "vitest";
import {
  AUTOCOLOR_LABEL,
  isAddinFormula,
  isExternalFormula,
  LINK_TOKENS_LABEL,
  type ShareIssue,
  type ShareScan,
  shareReport,
  summarizeShare,
} from "./share";

const EMPTY: ShareScan = {
  sheets: [],
  externalLinks: [],
  addinFormulas: [],
  brokenNames: [],
  skippedSheets: [],
  overlaysPainted: [],
  linkTokens: false,
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

  it("lists in workbook order, then the state the file carries", () => {
    const report = shareReport({
      sheets: [{ name: "Scratch", visibility: "Hidden" }],
      externalLinks: [
        { sheet: "Model", address: "B4", formula: "=[a.xlsx]S!A1" },
      ],
      addinFormulas: [
        { sheet: "Model", address: "C4", formula: "=PLSFIX.ROUND(C3,0)" },
      ],
      brokenNames: ["Costs"],
      skippedSheets: ["Data"],
      overlaysPainted: ["The audit overlay"],
      linkTokens: true,
      autocolorOnEdit: true,
    });

    expect(report.map((issue) => issue.kind)).toEqual([
      "hiddenSheet",
      "externalLink",
      "addinFormula",
      "brokenName",
      "skippedSheet",
      "overlayPainted",
      "linkTokens",
      "autocolorOnEdit",
    ]);
  });

  // The three this release introduced: paint that travels with the file,
  // formulas only this add-in can evaluate, and the tokens in the settings.
  it("reports a painted overlay by name and says who cannot clear it", () => {
    expect(
      shareReport(
        scan({
          overlaysPainted: ["The audit overlay", "The linked-cell highlight"],
        }),
      ),
    ).toEqual([
      {
        kind: "overlayPainted",
        label:
          "The audit overlay is still painted over cells; the reader has no add-in to clear it",
      },
      {
        kind: "overlayPainted",
        label:
          "The linked-cell highlight is still painted over cells; the reader has no add-in to clear it",
      },
    ]);
  });

  it("reports a PLSFIX formula by cell, the way an external link is reported", () => {
    expect(
      shareReport(
        scan({
          addinFormulas: [
            { sheet: "Model", address: "D9", formula: "=PLSFIX.ROUND(D8,-3)" },
          ],
        }),
      ),
    ).toEqual([
      { kind: "addinFormula", label: "Model!D9: =PLSFIX.ROUND(D8,-3)" },
    ]);
  });

  it("reports link tokens as something to break, never to delete", () => {
    expect(shareReport(scan({ linkTokens: true }))).toEqual([
      { kind: "linkTokens", label: LINK_TOKENS_LABEL },
    ]);
    expect(shareReport(scan({ linkTokens: false }))).toEqual([]);
  });
});

describe("isAddinFormula", () => {
  it("catches this add-in's own functions, prefixed or not", () => {
    expect(isAddinFormula("=PLSFIX.ROUND(B4,0)")).toBe(true);
    expect(isAddinFormula("=_xlfn.PLSFIX.ROUNDSUM(B4:B9,0)")).toBe(true);
    expect(isAddinFormula("=SUM(A1)+PLSFIX.ROUND(B4,0)")).toBe(true);
  });

  it("leaves an ordinary formula and a plain value alone", () => {
    expect(isAddinFormula("=ROUND(B4,0)")).toBe(false);
    expect(isAddinFormula("=SUM(B4:B9)")).toBe(false);
    expect(isAddinFormula("PLSFIX.ROUND")).toBe(false);
    expect(isAddinFormula(42)).toBe(false);
    expect(isAddinFormula(null)).toBe(false);
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

  it("counts the cells that need the add-in", () => {
    expect(summarizeShare(issues("addinFormula", 2), 3)).toBe(
      "3 sheets reset to A1; 2 cells needing the add-in left",
    );
    expect(summarizeShare(issues("addinFormula", 1), 3)).toBe(
      "3 sheets reset to A1; 1 cell needing the add-in left",
    );
  });

  it("gives a painted overlay and the link registry a clause each", () => {
    expect(
      summarizeShare(
        [...issues("overlayPainted", 1), ...issues("linkTokens", 1)],
        2,
      ),
    ).toBe(
      "2 sheets reset to A1; an overlay is still painted; link tokens travel with the file",
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
