// 08.09 links proof, Excel half: opens the pane, generates the link key, adds a "Proof"
// sheet with a four-series column chart and a doughnut, exports two demo charts, both
// proof charts and a table, and reports every toast (the doughnut one carries the reason).
async (page, ctx, pages, shot, h) => {
  const excel = h.hostPage(ctx, "x");
  const log = [];
  const step = async (name, fn) => {
    try {
      log.push(name + ": " + (await fn()));
    } catch (e) {
      log.push(
        name +
          " ERR " +
          String(e.message || e)
            .split("\n")[0]
            .slice(0, 220),
      );
    }
  };
  const paneOf = async (host, re, tab, button) => {
    let f = await h.frameByUrl(host, re, 2);
    if (!f) {
      await h.clickIn(host, "tab", tab).catch(() => false);
      await host.waitForTimeout(1500);
      await h.clickIn(host, "button", button);
      await host.waitForTimeout(20000);
      f = await h.frameByUrl(host, re);
    }
    return f;
  };
  await excel.bringToFront();
  const xp = await paneOf(
    excel,
    /modelis\/taskpane/,
    /^pls,fix$/,
    /^Model Tools$/,
  );
  if (!xp) return log.concat(["no excel pane"]);
  const xrun = (fn, arg) =>
    xp.evaluate(
      async ([src, arg]) => {
        const f = new Function("return (" + src + ")")();
        return Excel.run((c) => f(c, arg));
      },
      [fn.toString(), arg],
    );
  const xtoast = () =>
    xp.evaluate(() =>
      (document.querySelector("#toast")?.textContent || "")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 200),
    );
  const xtab = async (name) => {
    await xp
      .getByRole("tab", { name: new RegExp("^" + name + "$") })
      .first()
      .click({ timeout: 20000 });
    await excel.waitForTimeout(1200);
  };
  const nameBox = async (addr) => {
    for (const f of excel.frames()) {
      const nb = f.getByRole("combobox", { name: /Name Box/i }).first();
      if (await nb.count().catch(() => 0)) {
        await nb.click();
        await nb.fill(addr);
        await nb.press("Enter");
        await excel.waitForTimeout(2500);
        return "ok";
      }
    }
    return "no name box";
  };
  await step(
    "excel pane",
    async () =>
      (await xp.evaluate(
        () => document.querySelector("#connection-status")?.textContent,
      )) +
      " " +
      (await xp.evaluate(
        () => (document.body.innerText.match(/v\d+\.\d+\.\d+/) || [""])[0],
      )),
  );
  await step("proof sheet + charts", () =>
    xrun(async (c) => {
      const wb = c.workbook;
      const old = wb.worksheets.getItemOrNullObject("Proof");
      old.load("isNullObject");
      await c.sync();
      if (!old.isNullObject) {
        old.delete();
        await c.sync();
      }
      const ws = wb.worksheets.add("Proof");
      ws.getRange("A1:E6").values = [
        ["", "North", "South", "East", "West"],
        ["FY22", 120, 80, 60, 40],
        ["FY23", 130, 85, 70, 45],
        ["FY24", 150, 90, 75, 50],
        ["FY25", 160, 100, 80, 55],
        ["FY26", 175, 110, 90, 60],
      ];
      const col = ws.charts.add(
        "ColumnClustered",
        ws.getRange("A1:E6"),
        "Auto",
      );
      col.name = "Four series";
      col.title.text = "Four series";
      col.setPosition("G2", "N18");
      ws.getRange("A9:B14").values = [
        ["Slice", "Share"],
        ["A", 30],
        ["B", 25],
        ["C", 20],
        ["D", 15],
        ["E", 10],
      ];
      const d = ws.charts.add("Doughnut", ws.getRange("A9:B14"), "Auto");
      d.name = "Doughnut";
      d.title.text = "Doughnut";
      d.setPosition("G20", "N36");
      await c.sync();
      ws.charts.load("items/name");
      await c.sync();
      return ws.charts.items.map((x) => x.name).join(", ");
    }),
  );
  await step("links tab + key", async () => {
    await xtab("Links");
    await excel.waitForTimeout(1500);
    const gen = xp.locator("#generate-key");
    const visible = await gen.isVisible().catch(() => false);
    if (visible) {
      await gen.click({ timeout: 8000 });
      await excel.waitForTimeout(3000);
    }
    await xp
      .locator("#reveal-key")
      .click({ timeout: 5000 })
      .catch(() => {});
    await excel.waitForTimeout(800);
    const shown = await xp.evaluate(() =>
      (
        document.querySelector("#workspace-key-display")?.textContent || ""
      ).trim(),
    );
    return (
      "generate visible=" +
      visible +
      " key=" +
      shown.slice(0, 12) +
      "... toast=" +
      (await xtoast())
    );
  });
  const exportChart = async (sheet, chart) => {
    log.push("  nav " + sheet + ": " + (await nameBox(sheet + "!A1")));
    await xtab("Links");
    await excel.waitForTimeout(3500);
    const opts = await xp
      .locator("#export-chart-pick option")
      .allTextContents();
    await xp
      .locator("#export-chart-pick")
      .selectOption({ label: chart }, { timeout: 15000 });
    await xp.locator("#export-chart").click({ timeout: 15000 });
    for (let w = 0; w < 12; w++) {
      await excel.waitForTimeout(3000);
      const t = await xtoast();
      if (/Sent to PowerPoint|could not|refused|Error|failed/i.test(t))
        return t + "  [picker: " + opts.join(" | ") + "]";
    }
    return "no toast: " + (await xtoast());
  };
  await step("export Revenue chart", () =>
    exportChart("Bridge", "Revenue chart"),
  );
  await step("export Segment pie", () =>
    exportChart("Rounding", "Segment pie"),
  );
  await step("export Four series", () => exportChart("Proof", "Four series"));
  await step("export Doughnut", () => exportChart("Proof", "Doughnut"));
  await step("export table P&L!B4:E9", async () => {
    log.push("  nav P&L: " + (await nameBox("'P&L'!B4:E9")));
    await xtab("Links");
    await excel.waitForTimeout(1500);
    await xp.locator("#export-table").click({ timeout: 15000 });
    for (let w = 0; w < 12; w++) {
      await excel.waitForTimeout(3000);
      const t = await xtoast();
      if (/Sent to PowerPoint|Select|could not|Error|failed/i.test(t)) return t;
    }
    return "no toast: " + (await xtoast());
  });
  await step("workbook links", async () => {
    await xp
      .locator("#tab-links")
      .click()
      .catch(() => {});
    await excel.waitForTimeout(2000);
    return (
      (await xp.evaluate(() =>
        document
          .querySelector("#workbook-links")
          ?.innerText.replace(/\s+/g, " ")
          .slice(0, 300),
      )) || "(none)"
    );
  });
  await shot("proof-a-excel");
  return log;
};
