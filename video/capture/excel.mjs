// excel.mjs: Excel for the web specifics: activate a sheet and select a range through
// Office.js inside the pane frame, and turn cell addresses into capture rects from the
// Range geometry (points; x 4/3 at 100 % zoom) plus the grid origin read off Excel's own
// header strips. Invariant: a select starts from A1, so the sheet is scrolled to its corner
// and the geometry needs no scroll offset.
import { frameByUrl, rect } from "./office.mjs";

const PX_PER_PT = 4 / 3;

export async function select(frame, sheet, address) {
  await frame.evaluate(
    ({ sheet, address }) =>
      Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(sheet);
        ws.activate();
        ws.getRange("A1").select();
        await ctx.sync();
        ws.getRange(address).select();
        await ctx.sync();
      }),
    { sheet, address },
  );
  // the inspector follows Excel's selection events; an API select needs its refresh button
  const want = address.split(",")[0];
  for (let i = 0; i < 6; i++) {
    await frame.evaluate(() => document.getElementById("refresh-selection").click());
    await frame.page().waitForTimeout(1200);
    const shown = await frame.evaluate(() => document.getElementById("selection-address").textContent);
    if (shown.includes(want)) return;
  }
  throw new Error(`select ${sheet}!${address}: the inspector never showed it`);
}

// Moves a chart (points), as a modeller drags one into place.
export async function moveChart(frame, sheet, name, box) {
  await frame.evaluate(
    ({ sheet, name, box }) =>
      Excel.run(async (ctx) => {
        const c = ctx.workbook.worksheets.getItem(sheet).charts.getItem(name);
        Object.assign(c, box);
        await ctx.sync();
      }),
    { sheet, name, box },
  );
  await frame.page().waitForTimeout(1500);
}

// A chart's box in points, to put it back after it stepped aside.
export async function chartBox(frame, sheet, name) {
  return frame.evaluate(
    ({ sheet, name }) =>
      Excel.run(async (ctx) => {
        const c = ctx.workbook.worksheets.getItem(sheet).charts.getItem(name);
        c.load("left,top,width,height");
        await ctx.sync();
        return { left: c.left, top: c.top, width: c.width, height: c.height };
      }),
    { sheet, name },
  );
}

// Writes values or formulas into a range (the demo's own edits, like a new assumption).
export async function write(frame, sheet, address, values) {
  await frame.evaluate(
    ({ sheet, address, values }) =>
      Excel.run(async (ctx) => {
        ctx.workbook.worksheets.getItem(sheet).getRange(address).values = values;
        await ctx.sync();
      }),
    { sheet, address, values },
  );
  await frame.page().waitForTimeout(1500);
}

// Where cell A1's top-left corner sits on the page (CSS px).
async function gridOrigin(page) {
  const xf = await frameByUrl(page, /xlviewerinternal/);
  const fb = await (await xf.frameElement()).boundingBox();
  const o = await xf.evaluate(() => {
    const col = document.querySelector("[id$='columnHeadersDiv']");
    const row = document.querySelector("[id$='rowHeadersDiv']");
    if (!col || !row) throw new Error("grid: header strips not found");
    return { x: row.getBoundingClientRect().right, y: col.getBoundingClientRect().bottom };
  });
  return { x: fb.x + o.x, y: fb.y + o.y };
}

// specs: [[key, address]]; answers { key: [x, y, w, h] } in capture px.
export async function cellRects(page, frame, sheet, specs) {
  const origin = await gridOrigin(page);
  const geo = await frame.evaluate(
    ({ sheet, addrs }) =>
      Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(sheet);
        const rs = addrs.map((a) => {
          const r = ws.getRange(a);
          r.load("left,top,width,height");
          return r;
        });
        await ctx.sync();
        return rs.map((r) => [r.left, r.top, r.width, r.height]);
      }),
    { sheet, addrs: specs.map((s) => s[1]) },
  );
  const out = {};
  specs.forEach(([key], i) => {
    const [l, t, w, h] = geo[i];
    out[key] = rect({ x: origin.x + l * PX_PER_PT, y: origin.y + t * PX_PER_PT, width: w * PX_PER_PT, height: h * PX_PER_PT });
  });
  return out;
}

// The page box of a chart on the sheet (its top-left and size in points, like a range).
export async function chartRect(page, frame, sheet, name) {
  const origin = await gridOrigin(page);
  const g = await frame.evaluate(
    ({ sheet, name }) =>
      Excel.run(async (ctx) => {
        const c = ctx.workbook.worksheets.getItem(sheet).charts.getItem(name);
        c.load("left,top,width,height");
        await ctx.sync();
        return [c.left, c.top, c.width, c.height];
      }),
    { sheet, name },
  );
  return rect({ x: origin.x + g[0] * PX_PER_PT, y: origin.y + g[1] * PX_PER_PT, width: g[2] * PX_PER_PT, height: g[3] * PX_PER_PT });
}

// Every chart on a sheet by name, newest last (a tool's new chart is the one not seen before).
export async function chartNames(frame, sheet) {
  return frame.evaluate(
    (sheet) =>
      Excel.run(async (ctx) => {
        const cs = ctx.workbook.worksheets.getItem(sheet).charts;
        cs.load("items/name");
        await ctx.sync();
        return cs.items.map((c) => c.name);
      }),
    sheet,
  );
}
