// Proof of the dense-chart mitigation: a slot too narrow for its widest value
// label drops value labels for the whole column chart rather than wrap them;
// one too narrow for its widest category name thins the axis to every k-th
// label instead of an ellipsis wall; plus the recalibrated textWidth both rest on.

import { describe, expect, it } from "vitest";
import type { ChartData } from "./link/chart-model";
import { LABEL_SIZE, layoutChart, textWidth, type Text } from "./chart-shapes";

const base = {
  v: 1 as const,
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: "#14213D",
  title: null,
};

function texts(out: ReturnType<typeof layoutChart>): Text[] {
  return out.filter((one): one is Text => one.kind === "text");
}

// The same shape columnChart(points) in test/ppt.support.ts builds: a year
// per category, a three-digit-ish value climbing with it.
function columns(points: number): ChartData {
  const categories = Array.from({ length: points }, (_, i) => String(2021 + i));
  const values = categories.map((_, i) => 120 + 20 * i);
  return {
    ...base,
    kind: "column",
    categories,
    series: [
      {
        name: "Revenue",
        values,
        labels: values.map(String),
        colors: categories.map(() => "#B27E54"),
      },
    ],
  };
}

const WIDE_BOX = { left: 0, top: 0, width: 900, height: 300 };

describe("columnBars value labels", () => {
  it("draws none of the forty at 900 pt: the slot cannot hold a three-digit label", () => {
    const out = layoutChart(columns(40), WIDE_BOX);
    const values = texts(out).filter((one) => one.name.startsWith("label "));
    expect(values).toEqual([]);
    // The bars themselves are unaffected: forty still stand.
    expect(out.filter((one) => one.name.startsWith("bar ")).length).toBe(40);
  });

  it("keeps all six at 900 pt: the slot is not remotely dense", () => {
    const out = layoutChart(columns(6), WIDE_BOX);
    const values = texts(out).filter((one) => one.name.startsWith("label "));
    expect(values).toHaveLength(6);
    expect(values.map((one) => one.text)).toEqual(columns(6).series[0]!.labels);
  });
});

describe("categoryLabels thinning", () => {
  it("keeps every second or third year at forty columns in 900 pt", () => {
    const out = layoutChart(columns(40), WIDE_BOX);
    const categories = texts(out).filter((one) =>
      one.name.startsWith("category "),
    );
    // A slot of 22.5 pt cannot hold a four-character year (2021.."2060") un-
    // truncated: thinning kicks in rather than forty ellipses.
    expect(categories.length).toBeGreaterThan(0);
    expect(categories.length).toBeLessThan(40);
    expect(categories.some((one) => one.name === "category 0")).toBe(true);
    // Every shown label reads the real year, not an ellipsis: the merged
    // slots gave it the room a single one never had.
    expect(categories.every((one) => /^20\d\d$/.test(one.text))).toBe(true);
  });

  it("keeps all six categories at 900 pt, exactly as before the fix", () => {
    const out = layoutChart(columns(6), WIDE_BOX);
    const categories = texts(out).filter((one) =>
      one.name.startsWith("category "),
    );
    expect(categories.map((one) => one.text)).toEqual(columns(6).categories);
  });

  it("thins a bar chart's row labels the same way, by row height rather than width", () => {
    const rows = 30;
    const tall: ChartData = {
      ...base,
      kind: "bar",
      categories: Array.from({ length: rows }, (_unused, i) => `Row ${i}`),
      series: [
        {
          name: "s",
          values: Array.from({ length: rows }, (_unused, i) => 10 + i),
          labels: Array.from({ length: rows }, (_unused, i) => String(10 + i)),
          colors: Array.from({ length: rows }, () => "#2EC4B6"),
        },
      ],
    };
    // 200 pt of plot over thirty rows is a 6.67 pt row: under LABEL_HEIGHT
    // (18), so the row labels thin even though nothing here is wide.
    const out = layoutChart(tall, { left: 0, top: 0, width: 480, height: 200 });
    const categories = texts(out).filter((one) =>
      one.name.startsWith("category "),
    );
    expect(categories.length).toBeGreaterThan(0);
    expect(categories.length).toBeLessThan(rows);
    expect(categories.some((one) => one.name === "category 0")).toBe(true);
    // Bar value labels are a different mechanism (barLabel clamps them past
    // or inside the bar end) and are never dropped by this mitigation.
    const values = texts(out).filter((one) => one.name.startsWith("label "));
    expect(values).toHaveLength(rows);
  });
});

describe("textWidth after the wrap-fix recalibration", () => {
  it("covers PowerPoint's own text-box insets, not only the glyphs", () => {
    // Two 7.2 pt insets (PowerPoint's 0.1 in default on each side), rounded
    // up to 15, replace the old, too-thin 7 pt guess that let "1 519" wrap;
    // the per-character factor also grew from 0.55 em to 0.6 em.
    expect(textWidth("", LABEL_SIZE)).toBe(15);
    expect(textWidth("1 519", LABEL_SIZE)).toBeCloseTo(
      0.6 * LABEL_SIZE * 5 + 15,
      6,
    );
  });
});
