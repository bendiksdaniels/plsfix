// The model check's rules, away from Office.js: which formulas read as
// volatile or hardcoded, the order a report lists findings in, what the cap
// does to a long one, and the two strings the pane shows.

import { describe, expect, it } from "vitest";

import {
  buildReport,
  type Finding,
  hardcodeIn,
  KIND_LABELS,
  MODEL_CHECK_FINDING_CAP,
  type ModelCheckReport,
  reportText,
  summarize,
  volatileIn,
} from "./model-check";

function finding(
  kind: Finding["kind"],
  sheet: string | null,
  ref: string | null,
  note = "",
): Finding {
  return { kind, sheet, ref, count: 1, note };
}

function report(
  findings: Finding[],
  sheetOrder: string[] = ["Model", "Data"],
  skipped: string[] = [],
): ModelCheckReport {
  return buildReport({
    findings,
    sheetOrder,
    scanned: { sheets: sheetOrder.length, cells: 100 },
    skipped,
  });
}

describe("volatileIn", () => {
  it("names the first volatile function called, whatever its casing", () => {
    expect(volatileIn("=SUM(A1:A9)+offset(B1,1,0)")).toBe("OFFSET");
    expect(volatileIn('=Indirect("A"&B1)')).toBe("INDIRECT");
    expect(volatileIn("=TODAY()-A1")).toBe("TODAY");
  });

  it("reads the whole name, so RANDBETWEEN is not RAND", () => {
    expect(volatileIn("=RANDBETWEEN(1,9)")).toBe("RANDBETWEEN");
    expect(volatileIn("=RAND()")).toBe("RAND");
  });

  it("takes the leftmost of several", () => {
    expect(volatileIn("=NOW()+OFFSET(A1,1,1)")).toBe("NOW");
    expect(volatileIn("=OFFSET(A1,1,1)+NOW()")).toBe("OFFSET");
  });

  it("wants a call, not a word that looks like one", () => {
    expect(volatileIn("=MYOFFSET(A1,1,1)")).toBeNull();
    expect(volatileIn("=A1+OFFSET_HELPER")).toBeNull();
    expect(volatileIn("=SUM(A1:A9)")).toBeNull();
  });

  it("looks through the prefix Excel writes on an unresolved function", () => {
    expect(volatileIn("=_xlfn.INDIRECT(A1)")).toBe("INDIRECT");
  });

  // Excel stores a custom function's own formula under a second prefix,
  // _xldudf_<namespace>_ (underscores, no dot), once the workbook is saved.
  // That prefix is cut off the same way, so a custom function is judged by
  // its own name rather than by the opaque prefixed token.
  it("looks through the prefix Excel writes on a saved custom function", () => {
    expect(volatileIn("=_xldudf_NOW(A1)")).toBe("NOW");
    expect(volatileIn("=_xldudf_PLSFIX_CAGR(B11,G11,5)")).toBeNull();
  });
});

describe("hardcodeIn", () => {
  it("flags a number typed into a formula, on any sheet", () => {
    expect(hardcodeIn("=B2*1.1")).toBe(true);
    expect(hardcodeIn("=Data!B2*1.1")).toBe(true);
    expect(hardcodeIn("=[Budget.xlsx]Model!$B$4*2")).toBe(true);
  });

  it("leaves references, function names and quoted text alone", () => {
    expect(hardcodeIn("=SUM(A1:A10)")).toBe(false);
    expect(hardcodeIn("=LOG10(B4)")).toBe(false);
    expect(hardcodeIn('=IF(A1="2026",B1,C1)')).toBe(false);
  });

  it("says no to anything that is not a formula", () => {
    expect(hardcodeIn("1.1")).toBe(false);
    expect(hardcodeIn("")).toBe(false);
  });
});

describe("buildReport ordering", () => {
  it("lists sheets in workbook order, errors first, then reading order", () => {
    const built = report([
      finding("hardcodeInFormula", "Data", "A1"),
      finding("hardcodeInFormula", "Model", "A10"),
      finding("hardcodeInFormula", "Model", "A9"),
      finding("formulaError", "Model", "Z100"),
    ]);

    expect(
      built.findings.map(
        (entry) => `${entry.sheet ?? "-"}!${entry.ref ?? "-"}`,
      ),
    ).toEqual(["Model!Z100", "Model!A9", "Model!A10", "Data!A1"]);
  });

  it("puts a finding with no cell ahead of its sheet's cells", () => {
    const built = report([
      finding("hardcodeInFormula", "Model", "B2"),
      finding("hiddenSheet", "Model", null),
    ]);

    expect(built.findings.map((entry) => entry.kind)).toEqual([
      "hiddenSheet",
      "hardcodeInFormula",
    ]);
  });

  it("puts the workbook's own findings after every sheet", () => {
    const built = report([
      finding("brokenName", null, null, "Revenue"),
      finding("unusedStyle", null, null, "Header 2"),
      finding("hardcodeInFormula", "Data", "A1"),
    ]);

    expect(built.findings.map((entry) => entry.sheet)).toEqual([
      "Data",
      null,
      null,
    ]);
  });

  it("keeps a finding on a sheet the order does not name", () => {
    const built = report([finding("hiddenSheet", "Ghost", null)], ["Model"]);
    expect(built.findings).toHaveLength(1);
  });
});

describe("the finding cap", () => {
  it("passes a short list through untruncated", () => {
    const built = report([finding("formulaError", "Model", "A1", "#REF!")]);
    expect(built.truncated).toBe(false);
    expect(built.findings).toHaveLength(1);
  });

  it("cuts a long list to the cap and flags it", () => {
    const many = Array.from({ length: MODEL_CHECK_FINDING_CAP + 7 }, (_, i) =>
      finding("hardcodeInFormula", "Model", `A${String(i + 1)}`),
    );

    const built = report(many);

    expect(built.findings).toHaveLength(MODEL_CHECK_FINDING_CAP);
    expect(built.truncated).toBe(true);
    // The cap keeps the first page of the ordered list, not a random slice.
    expect(built.findings[0]?.ref).toBe("A1");
  });
});

describe("summarize", () => {
  it("counts the findings and the sheets", () => {
    const built = report([
      finding("formulaError", "Model", "A1"),
      finding("hardcodeInFormula", "Model", "B2"),
    ]);
    expect(summarize(built)).toBe("Model check: 2 findings on 2 sheets");
  });

  it("says so when there is nothing to flag", () => {
    expect(summarize(report([], ["Model"]))).toBe(
      "Model check: nothing to flag on 1 sheet",
    );
  });

  it("names how many sheets it could not read", () => {
    const built = report([], ["Model", "Data"], ["Big", "Bigger"]);
    expect(summarize(built)).toBe(
      "Model check: nothing to flag on 2 sheets (2 sheets skipped)",
    );
  });

  it("says the list was capped, and says both when both happened", () => {
    const many = Array.from({ length: MODEL_CHECK_FINDING_CAP + 1 }, (_, i) =>
      finding("hardcodeInFormula", "Model", `A${String(i + 1)}`),
    );

    expect(summarize(report(many))).toBe(
      "Model check: 500 findings on 2 sheets (list capped at 500)",
    );
    expect(summarize(report(many, ["Model", "Data"], ["Big"]))).toBe(
      "Model check: 500 findings on 2 sheets (list capped at 500, 1 sheet skipped)",
    );
  });
});

describe("reportText", () => {
  it("writes one line per finding: where, what, and the detail", () => {
    const built = report([
      finding("formulaError", "Model", "B4", "#REF!"),
      finding("hiddenSheet", "Data", null, "Hidden"),
      finding("brokenName", null, null, "Revenue"),
    ]);

    expect(reportText(built).split("\n")).toEqual([
      "Model!B4  Formula error  #REF!",
      "Data  Hidden sheet  Hidden",
      "Workbook  Broken name  Revenue",
    ]);
  });

  it("is empty when nothing was flagged", () => {
    expect(reportText(report([]))).toBe("");
  });
});

describe("KIND_LABELS", () => {
  it("labels every kind in short plain English", () => {
    for (const label of Object.values(KIND_LABELS)) {
      expect(label.length).toBeLessThanOrEqual(24);
      expect(label).not.toContain("—");
    }
  });
});
