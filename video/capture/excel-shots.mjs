// excel-shots.mjs: the Excel states the video is built from, in the order they happen on the
// fresh demo workbook: the P&L hero (Autocolor, Audit overlay, Precedents, Model check), then
// one before/after per toolbox chapter. Every press is the pane's real button; every write
// persists, so the list runs top to bottom once. Rect keys are what the scenes read.
import { paneRects, pin, press, ribbonBox, rect, scrollPane, setField, tab, unpin } from "./office.mjs";
import { cellRects, chartBox, chartNames, chartRect, moveChart, select } from "./excel.mjs";

const PANE_TOP = (f) => f.evaluate(() => { document.querySelector(".app-shell").scrollTop = 0; });

// One ring around several rows (the first `n` of a list): [x, y, w, h] in capture px.
function union(rects, n) {
  const rs = rects.slice(0, n);
  const x = Math.min(...rs.map((r) => r[0])), y = Math.min(...rs.map((r) => r[1]));
  const x1 = Math.max(...rs.map((r) => r[0] + r[2])), y1 = Math.max(...rs.map((r) => r[1] + r[3]));
  return [x, y, x1 - x, y1 - y];
}

// A "before" shows no toast from an earlier step (they time out by themselves a moment later).
const quiet = (frame) => unpin(frame);

// The pane's boot on the web ends with one warning toast (the custom-functions runtime never
// reports ready on a web sideload); wait for it, or 30 s, so it never lands in a shot.
async function settle(frame) {
  await unpin(frame);
  for (let i = 0; i < 60; i++) {
    const seen = await frame.evaluate(() =>
      window.__toastLog.some((t) => /never reported the add-in ready/.test(t.text)) ||
      /never reported the add-in ready/.test(document.getElementById("toast").innerText));
    if (seen) break;
    await frame.page().waitForTimeout(500);
  }
  await unpin(frame);
}

export { quiet };

export async function excelShots(page, f, snap) {
  const pane = async (specs) => paneRects(f, specs);
  const ribbonKey = async (name) => rect(await ribbonBox(page, "button", name));

  // Hero: the P&L model block, before and after Autocolor.
  await settle(f);
  await f.evaluate(() => document.getElementById("first-run-dismiss")?.click());
  await select(f, "P&L", "A10:H27");
  await PANE_TOP(f);
  const hero = await cellRects(page, f, "P&L", [["block", "A10:H27"], ["numbers", "C11:H27"], ["f16", "F16"], ["d17", "D17"], ["d14", "D14"], ["d16", "D16"], ["audit", "C11:H25"]]);
  const heroRibbon = { autocolor: await ribbonKey(/^Autocolor$/), audit: await ribbonKey(/^Audit overlay$/), precedents: await ribbonKey(/^Precedents$/) };
  await snap("x-pnl", { ...hero, ...heroRibbon, ...(await pane([["pane", ".app-shell"], ["tabs", "#tab-bar"]])) });
  const tColor = await press(page, f, '[data-action="autocolor"]');
  await PANE_TOP(f);
  await pin(f, tColor);
  await snap("x-pnl-color", { ...hero, toast: (await pane([["toast", "#toast"]])).toast });

  // Audit overlay over the numbers: the planted hardcode in F16 breaks its row.
  await quiet(f);
  await select(f, "P&L", "C11:H25");
  const tAudit = await press(page, f, '[data-action="audit-toggle"]');
  await PANE_TOP(f);
  await pin(f, tAudit);
  await snap("x-pnl-audit", { ...hero, toast: (await pane([["toast", "#toast"]])).toast });

  // Precedents of EBITDA 2025E: the trace panel lists them, the grid selects them.
  await quiet(f);
  await select(f, "P&L", "D17");
  await press(page, f, '[data-action="trace-precedents"]');
  await quiet(f);
  await scrollPane(f, "#audit-heading", 8);
  await snap("x-pnl-trace", { ...hero, ...(await pane([["trace", "#trace-panel"], ["chips", "#trace-chips button, #trace-chips .chip", true]])) });
  await press(page, f, '[data-action="audit-toggle"]');

  // Model check on the Workbook tab: the button before the press, then its findings.
  await quiet(f);
  await select(f, "P&L", "A1");
  await tab(f, "tab-workbook");
  await scrollPane(f, "#model-check-heading", 8);
  await snap("x-check-before", await pane([["run", '[data-action="run-model-check"]'], ["heading", "#model-check-heading"]]));
  await press(page, f, '[data-action="run-model-check"]', { wait: 120000 });
  await quiet(f);
  await scrollPane(f, "#model-check-heading", 8);
  await snap("x-check", await pane([["list", "#model-check-list"], ["heading", "#model-check-heading"], ["run", '[data-action="run-model-check"]'], ["rows", "#model-check-list > *", true]]));
  await tab(f, "tab-tools");

  // Chapter 1, formats: the wide grid's title row gets the brand Title preset (navy, white).
  await quiet(f);
  await select(f, "Data", "A8:G8");
  await scrollPane(f, "#format-heading", 8);
  const grid = await cellRects(page, f, "Data", [["grid", "A10:G15"], ["row", "A8:G8"]]);
  await snap("x-fmt", { ...grid, ...(await pane([["button", '[data-action="style-title"]'], ["presets", "#format-heading"]])) });
  const tTitle = await press(page, f, '[data-action="style-title"]');
  await scrollPane(f, "#format-heading", 8);
  await pin(f, tTitle);
  await snap("x-fmt-after", grid);

  // Chapter 2, charts: a waterfall from the bridge table. Both demo charts step aside first
  // and the new chart is dragged into the Revenue chart's place (the tool lands it below every
  // chart); afterwards the Revenue chart comes back for the PowerPoint act.
  await quiet(f);
  await moveChart(f, "Bridge", "EBITDA margin chart", { top: 700 });
  const revenueBox = await chartBox(f, "Bridge", "Revenue chart");
  await moveChart(f, "Bridge", "Revenue chart", { top: 720 });
  await select(f, "Bridge", "A11:B15");
  await scrollPane(f, "#charts-heading", 8);
  const bridgeBefore = await chartNames(f, "Bridge");
  const bridge = await cellRects(page, f, "Bridge", [["table", "A11:B15"]]);
  await snap("x-bridge", { ...bridge, ...(await pane([["charts", "#charts-heading"], ["waterfall", '[data-action="chart-waterfall"]']])) });
  const tFall = await press(page, f, '[data-action="chart-waterfall"]');
  const fall = (await chartNames(f, "Bridge")).find((n) => !bridgeBefore.includes(n));
  await moveChart(f, "Bridge", fall, revenueBox);
  await select(f, "Bridge", "A11:B15");
  await scrollPane(f, "#charts-heading", 8);
  await pin(f, tFall);
  await snap("x-bridge-after", { ...bridge, chart: await chartRect(page, f, "Bridge", fall) });
  await moveChart(f, "Bridge", fall, { top: 740 });
  await moveChart(f, "Bridge", "Revenue chart", revenueBox);

  // Chapter 3, templates: a DCF block written on a fresh sheet in the brand styles; its
  // columns are autofitted afterwards, as a modeller would.
  await quiet(f);
  await f.evaluate(() => Excel.run(async (ctx) => { ctx.workbook.worksheets.add("DCF"); await ctx.sync(); }));
  await select(f, "DCF", "B2");
  await scrollPane(f, "#templates-heading", 8);
  await snap("x-tpl", { ...(await cellRects(page, f, "DCF", [["cell", "B2"]])), ...(await pane([["dcf", '[data-action="template-dcf"]'], ["grid", "#templates-heading"]])) });
  const tTpl = await press(page, f, '[data-action="template-dcf"]');
  const block = await f.evaluate(() => Excel.run(async (ctx) => {
    const r = ctx.workbook.getSelectedRange();
    r.load("address");
    r.format.autofitColumns();
    await ctx.sync();
    return r.address.split("!")[1];
  }));
  await page.waitForTimeout(1500);
  await scrollPane(f, "#templates-heading", 8);
  await pin(f, tTpl);
  await snap("x-tpl-after", await cellRects(page, f, "DCF", [["cell", "B2"], ["block", block]]));

  // Chapter 4, Find a combination: which invoice lines add up to the target in B26.
  await quiet(f);
  const target = await f.evaluate(() => Excel.run(async (ctx) => { const r = ctx.workbook.worksheets.getItem("Variance").getRange("B26"); r.load("values"); await ctx.sync(); return r.values[0][0]; }));
  await select(f, "Variance", "B11:B24");
  await setField(f, "reconcile-target", target);
  await PANE_TOP(f);
  const lines = await cellRects(page, f, "Variance", [["list", "B11:B24"], ["target", "B26"]]);
  await snap("x-rec", { ...lines, ...(await pane([["find", '[data-action="reconcile-find"]'], ["input", "#reconcile-target"]])) });
  await press(page, f, '[data-action="reconcile-find"]');
  await quiet(f);
  const picked = await f.evaluate(() => Excel.run(async (ctx) => { const r = ctx.workbook.getSelectedRanges(); r.load("address"); await ctx.sync(); return r.address; }));
  await PANE_TOP(f);
  const hits = picked.split(",").map((a, i) => [`hit${i}`, a.split("!").pop()]);
  await snap("x-rec-after", { ...lines, ...(await cellRects(page, f, "Variance", hits)), ...(await pane([["result", "#reconcile-result"]])) });

  // Chapter 5, the whole file: Super Find across every sheet. "EBITDA 2024" sits in the
  // bridge table on screen and on the hidden Scratch sheet, which Excel's own Find skips.
  await quiet(f);
  await select(f, "Bridge", "A1");
  await tab(f, "tab-workbook");
  await setField(f, "find-query", "EBITDA 2024");
  await scrollPane(f, "#find-heading", 8);
  const found = await cellRects(page, f, "Bridge", [["cell", "A11"]]);
  await snap("x-find-before", { ...found, ...(await pane([["run", "#find-run"], ["query", "#find-query"]])) });
  await press(page, f, "#find-run");
  await quiet(f);
  await scrollPane(f, "#find-heading", 8);
  await snap("x-find", { ...found, ...(await pane([["results", "#find-results"], ["query", "#find-query"], ["hits", "#find-results > *", true], ["hidden", '#find-results > *:has-text("Scratch")']])) });
  await scrollPane(f, "#sheets-heading", 8);
  await snap("x-sheets", await pane([["explorer", "#sheets-heading"]]));

  // Chapter 6, brand: the palette, the house number style, the keyboard shortcuts.
  await tab(f, "tab-brand");
  await PANE_TOP(f);
  const brand = await pane([["preview", "#brand-preview"], ["palette", "#palette-list .palette-row", true]]);
  await snap("x-brand", { preview: brand.preview, palette: union(brand.palette, 2) });
  await scrollPane(f, "#shortcuts-heading", 8);
  const keys = await pane([["heading", "#shortcuts-heading"], ["rows", "#shortcuts-list > *", true]]);
  await snap("x-keys", { heading: keys.heading, keys: union(keys.rows, 3) });
  await tab(f, "tab-tools");
  await PANE_TOP(f);
}
