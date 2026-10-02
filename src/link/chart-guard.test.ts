import { describe, expect, it } from "vitest";
import type { ChartData } from "./chart-model";
import { CHART_UNREADABLE, guardChart } from "./chart-guard";
import { decodePayload, encodePayload, type PicturePayload } from "./model";

const src = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

const CHART: ChartData = {
  v: 1,
  kind: "column",
  title: "Revenue",
  categories: ["2024", "2025"],
  series: [
    {
      name: "Revenue",
      values: [120, 140],
      labels: ["120", "140"],
      colors: ["#B27E54", "#B27E54"],
    },
  ],
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: "#14213D",
};

function picture(extra: Record<string, unknown> = {}): PicturePayload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png: "iVBORw0KGgo=",
    src,
    pushedAt: "2026-09-12T08:00:00.000Z",
    hash: "0".repeat(64),
    ...extra,
  } as PicturePayload;
}

describe("guardChart", () => {
  it("hands a payload with no chart back untouched", () => {
    const payload = picture();
    expect(guardChart(payload)).toBe(payload);
  });

  it("hands a payload whose chart passes the validator back untouched", () => {
    const payload = picture({ chart: CHART });
    expect(guardChart(payload)).toBe(payload);
  });

  it("drops a chart the validator refuses and says why", () => {
    const guarded = guardChart(picture({ chart: { v: 1, kind: "column" } }));
    expect(guarded).toMatchObject({
      kind: "picture",
      chartIssue: CHART_UNREADABLE,
    });
    expect("chart" in (guarded as Record<string, unknown>)).toBe(false);
  });

  it("leaves the payload's own chartIssue standing", () => {
    const guarded = guardChart(
      picture({ chart: { nonsense: true }, chartIssue: "7 series" }),
    );
    expect(guarded).toMatchObject({ chartIssue: "7 series" });
    expect("chart" in (guarded as Record<string, unknown>)).toBe(false);
  });

  it("leaves a table, a text and a non-record alone", () => {
    const table = { v: 1, kind: "table", chart: { nonsense: true } };
    expect(guardChart(table)).toBe(table);
    const text = { v: 1, kind: "text", text: "12.4" };
    expect(guardChart(text)).toBe(text);
    expect(guardChart(null)).toBe(null);
    expect(guardChart("not a payload")).toBe("not a payload");
  });

  it("says the same sentence the panes show after 'as a picture'", () => {
    expect(CHART_UNREADABLE).toBe("chart data unreadable");
  });
});

describe("decodePayload over the guard", () => {
  it("keeps the picture when the chart beside it is unreadable", () => {
    const bytes = encodePayload(
      picture({ chart: { v: 1, kind: "column", categories: "nope" } }),
    );
    const decoded = decodePayload(bytes);
    expect(decoded.kind).toBe("picture");
    expect(decoded).toMatchObject({
      png: "iVBORw0KGgo=",
      chartIssue: CHART_UNREADABLE,
    });
    expect("chart" in decoded).toBe(false);
  });

  it("keeps a chart the validator accepts", () => {
    const decoded = decodePayload(encodePayload(picture({ chart: CHART })));
    expect(decoded).toMatchObject({ chart: CHART });
    expect("chartIssue" in decoded).toBe(false);
  });

  it("still refuses a payload that is not a payload at all", () => {
    const bytes = new TextEncoder().encode(JSON.stringify({ v: 1 }));
    expect(() => decodePayload(bytes)).toThrow("not a link payload");
  });
});
