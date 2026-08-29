// How the UX gate speaks: turns one measured combination into defect lines
// and prints the run summary. It owns the wording and the pass/fail count and
// nothing else - the caller decides the exit code from the total it returns.

const MIN_TAP_HEIGHT = 32;
const MIN_FONT_PX = 11;

const CLASSES = [
  ["overflow", "overflow"],
  ["scrollers", "horizontal-scroll"],
  ["tabStrip", "tab-strip"],
  ["shortControls", `short-control (<${String(MIN_TAP_HEIGHT)}px)`],
  ["clippedLabels", "clipped-label"],
  ["smallText", `small-text (<${String(MIN_FONT_PX)}px)`],
  ["toast", "toast-over-tabs"],
];

export function defectLines(combo) {
  const lines = [];
  for (const [key, label] of CLASSES) {
    for (const item of combo.defects[key] ?? []) {
      lines.push(`  ${label}: ${item}`);
    }
  }
  for (const stop of combo.invisibleFocus ?? []) {
    lines.push(`  invisible-focus: ${stop.desc}`);
  }
  return lines;
}

function title(combo) {
  return `=== ${combo.pane} / ${combo.state} / ${combo.tab} / ${String(combo.width)}px ===`;
}

export function printCombos(combos, log) {
  let total = 0;
  for (const combo of combos) {
    const lines = defectLines(combo);
    if (lines.length === 0) {
      log(`${title(combo)} clean (screenshot: ${combo.screenshot})`);
      continue;
    }
    log(
      `${title(combo)} ${String(lines.length)} defect(s) (screenshot: ${combo.screenshot})`,
    );
    for (const line of lines) log(line);
    total += lines.length;
  }
  return total;
}

export function printSummary(combos, total, log) {
  const sample = combos.find((combo) => (combo.focusOrder ?? []).length > 0);
  if (sample) {
    log("");
    log(
      `focus order sample (${sample.pane} / ${sample.tab} / ${String(sample.width)}px):`,
    );
    log(
      "  " +
        sample.focusOrder
          .map((stop) => stop.desc + (stop.visible ? "" : " [INVISIBLE]"))
          .join(" -> "),
    );
  }
  log("");
  log(
    total === 0
      ? `PASS - 0 defects across ${String(combos.length)} pane/state/tab/width combinations`
      : `FAIL - ${String(total)} defect(s) across ${String(combos.length)} combinations`,
  );
}
