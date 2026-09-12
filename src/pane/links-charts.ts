// The Links tab's chart picker: which chart a chart export targets when
// nothing is selected. Redrawn on tab-open and on a sheet change; an ExcelApi
// 1.7 host without the worksheet-activated event keeps the tab-open refresh.
// Office.js only reaches here through ../excel.

import { listActiveSheetCharts, watchActiveSheet } from "../excel";
import { ANCHOR_PREFIX } from "../link/model";

interface ChartPickTab {
  chartPick: HTMLSelectElement;
  deps: { root: ParentNode };
}

// A chart exported as a link is renamed to its anchor (PLSFIX_LINK_<id>), so
// the picker shows what the chart is instead of that bookkeeping name; the
// option's own value stays the real name, since an export by name has to keep
// finding it under whichever label is showing.
const LINK_LABEL_CHARS = 4;

function chartPickLabel(name: string): string {
  if (!name.startsWith(ANCHOR_PREFIX)) return name;
  return `Linked chart · ${name.slice(-LINK_LABEL_CHARS)}`;
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
    ...names.map((name) => new Option(chartPickLabel(name), name)),
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
