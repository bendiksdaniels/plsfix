// Ensures the Excel pane is visible (reopen via the ribbon when the shared
// runtime survives a closed pane), returns its frame box; optional tab shots.
async (page, ctx, pages, shot, h) => {
  const excel = h.hostPage(ctx, "x");
  const visible = async () => {
    const f = await h.frameByUrl(excel, /modelis\/taskpane/, 1);
    if (!f) return null;
    const el = await f.frameElement().catch(() => null);
    const box = el ? await el.boundingBox() : null;
    return box && box.width > 50 ? box : null;
  };
  let box = await visible();
  if (!box) {
    await h.clickIn(excel, "tab", /^pls,fix$/).catch(() => false);
    await excel.waitForTimeout(1500);
    await h.clickIn(excel, "button", /^Model Tools$/);
    await excel.waitForTimeout(12000);
    box = await visible();
  }
  if (!box) return "pane not visible";
  const xp = await h.frameByUrl(excel, /modelis\/taskpane/);
  const out = { box: [box.x, box.y, box.width, box.height].map(Math.round) };
  if (process.env.SHOTS) {
    for (const tab of ["Tools", "Workbook", "Links", "Brand"]) {
      const t = xp
        .getByRole("tab", { name: new RegExp("^" + tab + "$") })
        .first();
      if (await t.count()) await t.click({ timeout: 15000 });
      await excel.waitForTimeout(900);
      await excel.screenshot({
        path: h.shotsDir() + "/ui-" + tab.toLowerCase() + ".png",
        clip: {
          x: box.x,
          y: box.y,
          width: box.width,
          height: Math.min(box.height, 900),
        },
      });
    }
    const t = xp.getByRole("tab", { name: /^Tools$/ }).first();
    if (await t.count()) await t.click();
    out.shots = true;
  }
  return out;
};
