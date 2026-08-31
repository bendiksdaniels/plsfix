// The Links tab's chart picker: which chart a chart export targets when
// nothing is selected. Redrawn on tab-open and on a sheet change; an ExcelApi
// 1.7 host without the worksheet-activated event keeps the tab-open refresh.
// Office.js only reaches here through ../excel.

import { listActiveSheetCharts, watchActiveSheet } from "../excel";

interface ChartPickTab {
  chartPick: HTMLSelectElement;
  deps: { root: ParentNode };
}

// The chart list stays hidden on a sheet without charts; the first option
// keeps "whatever is selected" as the default.
export async function refreshChartPick(tab: ChartPickTab): Promise<void> {
  let names: string[] = [];
  try {
    names = await listActiveSheetCharts();
  } catch {
    names = [];
  }
  const current = Array.from(tab.chartPick.options)
    .map((option) => option.value)
    .filter(Boolean);
  if (
    current.join("\n") === names.join("\n") &&
    tab.chartPick.options.length > 0
  ) {
    return;
  }
  const keep = tab.chartPick.value;
  tab.chartPick.replaceChildren(
    new Option("Selected chart", ""),
    ...names.map((name) => new Option(name, name)),
  );
  tab.chartPick.value = names.includes(keep) ? keep : "";
  tab.chartPick.hidden = names.length === 0;
}

// The chart list belongs to the active sheet: it is redrawn when the tab is
// opened and when the modeller moves to another sheet. A host without the
// worksheet event (ExcelApi 1.7) keeps the tab-open refresh.
export function watchSheetChanges(tab: ChartPickTab): void {
  tab.deps.root
    .querySelector("#tab-links")
    ?.addEventListener("click", () => void refreshChartPick(tab));
  watchActiveSheet(() => refreshChartPick(tab));
}
