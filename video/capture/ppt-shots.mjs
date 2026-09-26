// ppt-shots.mjs: the Excel to PowerPoint link on both documents, in the order a modeller does
// it: export the Revenue chart, insert it on the demo deck's slide 3 (left half), change a
// growth assumption in Excel and Push all, PowerPoint sees "Update available", Update all,
// then the Tools tab lining up three shapes. Slide geometry comes from the editor's slide
// element (points x scale).
import { paneRects, pin, press, prime, rect, scrollPane, setField, tab, unpin } from "./office.mjs";
import { cellRects, chartNames, chartRect, select, write } from "./excel.mjs";

const CHART = "Revenue chart";
const LINK_SLIDE = 3; // "Choose the spot": the Revenue chart goes to its left half
const TOOL_SLIDE = 2; // an empty slide for the object tools

// The slide as it sits in the editor: its page box and CSS px per point.
async function slideBox(page) {
  for (const f of page.frames()) {
    const href = await f.evaluate(() => location.href).catch(() => "");
    if (!/ppt\.aspx/.test(href)) continue;
    // the editor's slide view; thumbnails use the same id prefix, so take the widest
    let b = null;
    for (const l of await f.locator("[id^='SlideRootViewElement']").all()) {
      const box = await l.boundingBox().catch(() => null);
      if (box && (!b || box.width > b.width)) b = box;
    }
    if (b && b.width > 400) return { ...b, k: b.width / 960 };
  }
  throw new Error("slide: the editor's slide element was not found");
}

async function goToSlide(frame, n) {
  await frame.evaluate((n) => PowerPoint.run(async (ctx) => {
    const s = ctx.presentation.slides.getItemAt(n - 1);
    s.load("id");
    await ctx.sync();
    ctx.presentation.setSelectedSlides([s.id]);
    await ctx.sync();
  }), n);
  await frame.page().waitForTimeout(2500);
}

// Every top-level shape on a slide: [name, left, top, width, height] in points.
async function shapes(frame, n) {
  return frame.evaluate((n) => PowerPoint.run(async (ctx) => {
    const list = ctx.presentation.slides.getItemAt(n - 1).shapes;
    list.load("items/name,items/left,items/top,items/width,items/height,items/id");
    await ctx.sync();
    return list.items.map((s) => [s.name, s.left, s.top, s.width, s.height, s.id]);
  }), n);
}

// PowerPoint reads the Inbox of the key it holds; the documented pairing pastes Excel's key into
// PowerPoint's Settings (the key moves between the two panes only, never to disk or the log).
async function pair(xf, pf) {
  const key = await xf.evaluate(() => OfficeRuntime.storage.getItem("plsfix.link.workspace.v1"));
  if (!key) throw new Error("pair: Excel holds no link key");
  const held = await pf.evaluate(() => localStorage.getItem("plsfix.link.workspace.v1"));
  if (held === key) return;
  await tab(pf, "tab-settings");
  await setField(pf, "workspace-key", key);
  await press(pf.page(), pf, "#save-key");
  console.log("capture: PowerPoint paired with Excel's link key");
}

const onPage = (sb, [, l, t, w, h]) => rect({ x: sb.x + l * sb.k, y: sb.y + t * sb.k, width: w * sb.k, height: h * sb.k });

export async function linkShots({ xl, xf, pp, snap, openPane }) {
  const sx = snap(xl), sp = snap(pp);

  // Excel: a fresh link key per run, so the Inbox holds only this run's export (Generate asks
  // for a second click while a key exists).
  await unpin(xf);
  await select(xf, "Bridge", "A1");
  await tab(xf, "tab-links");
  await xf.evaluate(() => document.getElementById("generate-key").click());
  await xl.waitForTimeout(500);
  const armed = await xf.evaluate(() => /again to confirm/i.test(document.getElementById("generate-key").textContent));
  if (armed) await press(xl, xf, "#generate-key"); // replacing a key asks twice, within 5 s
  else await xl.waitForTimeout(3000); // a first key is made on the first press

  // PowerPoint: paired with the fresh key, slide 3 in front, the Inbox still empty.
  const pf = await openPane(pp, "p", /^Links$/);
  await prime(pf);
  await pair(xf, pf);
  await pf.evaluate(() => document.getElementById("first-run-ppt-dismiss")?.click());
  await unpin(pf);
  await goToSlide(pf, LINK_SLIDE);
  await tab(pf, "tab-inbox");
  await setField(pf, "insert-slide", LINK_SLIDE);
  await setField(pf, "insert-where", "left-half");
  let sb = await slideBox(pp);
  await sp("p-empty", { slide: rect(sb), ...(await paneRects(pf, [["inbox", "#inbox-list"], ["where", "#insert-where"]])) });

  // Excel: the Revenue chart leaves through the Links tab.
  const names0 = await chartNames(xf, "Bridge");
  await setField(xf, "export-chart-pick", CHART);
  await scrollPane(xf, "#export-heading", 8);
  const chart = await chartRect(xl, xf, "Bridge", CHART);
  await sx("x-link", { chart, ...(await paneRects(xf, [["export", "#export-chart"], ["pick", "#export-chart-pick"]])) });
  const tExport = await press(xl, xf, "#export-chart");
  // a chart link is anchored by the chart's own name, so the export renames it
  const linked = (await chartNames(xf, "Bridge")).find((n) => !names0.includes(n)) ?? CHART;
  await scrollPane(xf, "#export-heading", 8);
  await pin(xf, tExport);
  await sx("x-link-after", { chart, ...(await paneRects(xf, [["export", "#export-chart"], ["toast", "#toast"]])) });

  // PowerPoint: the export waits in the Inbox; Insert puts it on slide 3's left half.
  await unpin(pf);
  await pf.evaluate(() => document.getElementById("refresh-inbox").click());
  await pf.waitForSelector(".inbox-row .inbox-insert", { timeout: 60000 });
  await unpin(pf);
  const before = (await shapes(pf, LINK_SLIDE)).map((s) => s[5]);
  sb = await slideBox(pp);
  await sp("p-inbox", { slide: rect(sb), ...(await paneRects(pf, [["row", ".inbox-row"], ["insert", ".inbox-row .inbox-insert"], ["where", "#insert-where"]])) });
  const tInsert = await press(pp, pf, ".inbox-row .inbox-insert", { wait: 120000 });
  const added = (await shapes(pf, LINK_SLIDE)).find((s) => !before.includes(s[5]));
  if (!added) throw new Error(`insert: nothing new on slide ${LINK_SLIDE} (${tInsert.text})`);
  await pin(pf, tInsert);
  sb = await slideBox(pp);
  await sp("p-inserted", { slide: rect(sb), chart: onPage(sb, added), ...(await paneRects(pf, [["linksTab", "#tab-links"]])) });

  // Excel: revenue growth for 2025E goes from 8 % to 25 % (a change the slide can show from
  // across the room), the chart follows, then Push all on the Links tab.
  await unpin(xf);
  await tab(xf, "tab-tools");
  await select(xf, "Assumptions", "C11");
  const cell = await cellRects(xl, xf, "Assumptions", [["cell", "C11"], ["row", "A11:G11"]]);
  await sx("x-assume", cell);
  await write(xf, "Assumptions", "C11", [[0.25]]);
  await select(xf, "Assumptions", "C11");
  await sx("x-assume-after", cell);
  await select(xf, "Bridge", "A1");
  await tab(xf, "tab-links");
  await scrollPane(xf, "#links-heading", 8);
  await sx("x-push-before", { chart: await chartRect(xl, xf, "Bridge", linked), ...(await paneRects(xf, [["push", "#push-all"]])) });
  const tPush = await press(xl, xf, "#push-all");
  await scrollPane(xf, "#links-heading", 8);
  await pin(xf, tPush);
  await sx("x-push", { chart: await chartRect(xl, xf, "Bridge", linked), ...(await paneRects(xf, [["push", "#push-all"], ["toast", "#toast"]])) });

  // PowerPoint: the Links list reads "Update available" once it has looked again (a push
  // takes a moment to arrive), then Update all repaints the chart where it stands.
  await unpin(pf);
  await tab(pf, "tab-links");
  let waiting = false;
  for (let i = 0; i < 8 && !waiting; i++) {
    await press(pp, pf, "#refresh-links");
    waiting = await pf.evaluate(() => !!document.querySelector("#link-rows .badge.update"));
    if (!waiting) await pp.waitForTimeout(5000);
  }
  if (!waiting) throw new Error("links: PowerPoint never saw the pushed update");
  await unpin(pf);
  await sp("p-links", { slide: rect(sb), chart: onPage(sb, added), ...(await paneRects(pf, [["update", "#update-all"], ["rows", "#link-rows"], ["status", "#link-rows .badge.update"]])) });
  const tUpdate = await press(pp, pf, "#update-all", { wait: 180000 });
  const redrawn = (await shapes(pf, LINK_SLIDE)).find((s) => !before.includes(s[5])) ?? added;
  await pin(pf, tUpdate);
  sb = await slideBox(pp);
  await sp("p-updated", { slide: rect(sb), chart: onPage(sb, redrawn), ...(await paneRects(pf, [["update", "#update-all"]])) });

  // PowerPoint Tools: three shapes drawn loose, then lined up and spaced evenly.
  await unpin(pf);
  await goToSlide(pf, TOOL_SLIDE);
  const ids = await pf.evaluate((n) => PowerPoint.run(async (ctx) => {
    const slide = ctx.presentation.slides.getItemAt(n - 1);
    const list = slide.shapes;
    // the slide's empty body placeholder ("Click to add text") goes first, as a person tidies it
    list.load("items/type,items/name");
    await ctx.sync();
    const holders = list.items.filter((s) => s.type === "Placeholder");
    holders.forEach((s) => s.textFrame.textRange.load("text"));
    await ctx.sync();
    holders.filter((s) => !s.textFrame.textRange.text.trim()).forEach((s) => s.delete());
    await ctx.sync();
    const spots = [[110, 190, "#14213D"], [390, 290, "#2EC4B6"], [640, 160, "#8A94A6"]];
    const made = spots.map(([left, top, fill]) => {
      const s = list.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, { left, top, width: 180, height: 110 });
      s.fill.setSolidColor(fill);
      s.lineFormat.visible = false;
      return s;
    });
    made.forEach((s) => s.load("id"));
    await ctx.sync();
    slide.setSelectedShapes(made.map((s) => s.id));
    await ctx.sync();
    return made.map((s) => s.id);
  }), TOOL_SLIDE);
  await tab(pf, "tab-tools");
  await setField(pf, "object-align-mode", "middle");
  await setField(pf, "object-distribute-axis", "horizontal");
  sb = await slideBox(pp);
  const boxes = async () => Object.fromEntries((await shapes(pf, TOOL_SLIDE)).filter((s) => ids.includes(s[5])).map((s, i) => [`shape${i}`, onPage(sb, s)]));
  await sp("p-tools", { slide: rect(sb), ...(await boxes()), ...(await paneRects(pf, [["align", "#align-objects"], ["distribute", "#distribute-objects"]])) });
  await press(pp, pf, "#align-objects");
  const tDist = await press(pp, pf, "#distribute-objects");
  await pin(pf, tDist);
  await sp("p-tools-after", { slide: rect(sb), ...(await boxes()) });
}
