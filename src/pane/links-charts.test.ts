// @vitest-environment jsdom
// The chart picker's option labels: a modeller's own chart name shows as is,
// but a chart already exported as a link was renamed to its anchor
// (PLSFIX_LINK_<id>), so the picker shows what it is instead of the
// bookkeeping name. The option's value stays the real name either way, since
// an export by name has to keep finding it.

import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../excel", () => ({
  listActiveSheetCharts: vi.fn(async () => []),
  watchActiveSheet: vi.fn(),
}));

import { listActiveSheetCharts } from "../excel";
import { refreshChartPick } from "./links-charts";

function chartPick(): HTMLSelectElement {
  document.body.innerHTML = "<select id='chart-pick'></select>";
  return document.getElementById("chart-pick") as HTMLSelectElement;
}

describe("refreshChartPick", () => {
  beforeEach(() => {
    vi.mocked(listActiveSheetCharts).mockReset();
  });

  it("shows a modeller's own chart name as is", async () => {
    vi.mocked(listActiveSheetCharts).mockResolvedValue(["Revenue chart"]);
    const select = chartPick();

    await refreshChartPick({ chartPick: select, deps: { root: document } });

    expect(select.options[1]?.textContent).toBe("Revenue chart");
    expect(select.options[1]?.value).toBe("Revenue chart");
  });

  it("labels a linked chart by its last four characters, value unchanged", async () => {
    vi.mocked(listActiveSheetCharts).mockResolvedValue([
      "PLSFIX_LINK_a1b2c3d4",
    ]);
    const select = chartPick();

    await refreshChartPick({ chartPick: select, deps: { root: document } });

    expect(select.options[1]?.textContent).toBe("Linked chart · c3d4");
    expect(select.options[1]?.value).toBe("PLSFIX_LINK_a1b2c3d4");
  });

  it("labels each linked chart in a mixed list, own names untouched", async () => {
    vi.mocked(listActiveSheetCharts).mockResolvedValue([
      "PLSFIX_LINK_deadbeef",
      "Margin bridge",
    ]);
    const select = chartPick();

    await refreshChartPick({ chartPick: select, deps: { root: document } });

    expect(
      Array.from(select.options).map((option) => option.textContent),
    ).toEqual(["Selected chart", "Linked chart · beef", "Margin bridge"]);
  });
});
