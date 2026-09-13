// Task 8 web proof: three chart exports from Excel for the web, inserts on slide 3 of the
// PowerPoint deck, read-back of groups vs pictures, then a source change + Push all + Update all.
async (page, ctx, pages, shot, h) => {
  const out = h.shotsDir() + "/";
  const excel = h.hostPage(ctx, "x"),
    ppt = h.hostPage(ctx, "p");
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
            .slice(0, 200),
      );
    }
  };
  const paneOf = async (host, re, tab, button) => {
    let f = await h.frameByUrl(host, re, 2);
    if (f) {
      await f.evaluate(() => location.reload());
      await host.waitForTimeout(18000);
      f = await h.frameByUrl(host, re);
    }
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
  if (!xp) return "no excel pane";
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
        .slice(0, 140),
    );
  const xtab = async (name) => {
    const t = xp
      .getByRole("tab", { name: new RegExp("^" + name + "$") })
      .first();
    await t.click({ timeout: 20000 });
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
        return;
      }
    }
  };
  const exportChart = async (sheet, chart) => {
    await nameBox(sheet + "!A1");
    await xtab("Links");
    await excel.waitForTimeout(3000);
    await xp
      .locator("#export-chart-pick")
      .selectOption(chart, { timeout: 15000 });
    await xp.locator("#export-chart").click({ timeout: 15000 });
    await excel.waitForTimeout(16000);
    return await xtoast();
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
  await step("export Revenue chart (column)", () =>
    exportChart("Bridge", "Revenue chart"),
  );
  await step("export Segment pie", () =>
    exportChart("Rounding", "Segment pie"),
  );
  await step("export EBITDA margin chart (line)", () =>
    exportChart("Bridge", "EBITDA margin chart"),
  );
  await ppt.bringToFront();
  const pp = await paneOf(ppt, /modelis\/pptpane/, /^pls,fix$/, /^Links$/);
  if (!pp) return log.concat(["no ppt pane"]);
  const prun = (fn, arg) =>
    pp.evaluate(
      async ([src, arg]) => {
        const f = new Function("return (" + src + ")")();
        return Promise.race([
          PowerPoint.run((c) => f(c, arg)),
          new Promise((_, rej) =>
            setTimeout(() => rej(new Error("timeout 120s")), 120000),
          ),
        ]);
      },
      [fn.toString(), arg],
    );
  const ptoast = () =>
    pp.evaluate(() =>
      (document.querySelector("#toast")?.textContent || "")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 160),
    );
  const slide3 = () =>
    prun(async (c) => {
      const shapes = c.presentation.slides.getItemAt(2).shapes;
      shapes.load(
        "items/id,items/name,items/type,items/left,items/top,items/width,items/height",
      );
      await c.sync();
      const out = [];
      for (const s of shapes.items) {
        if (s.type === "Placeholder") continue;
        const tags = s.tags;
        tags.load("items/key");
        let kids = null;
        if (s.type === "Group") {
          kids = s.group.shapes;
          kids.load("items/name,items/type");
        }
        await c.sync();
        out.push(
          s.type +
            " '" +
            s.name.slice(0, 40) +
            "' @" +
            [s.left, s.top, s.width, s.height].map(Math.round).join(",") +
            " tags " +
            tags.items.map((t) => t.key).join("/") +
            (kids
              ? " kids " +
                kids.items.length +
                " (" +
                [...new Set(kids.items.map((k) => k.type))].join("/") +
                ")"
              : ""),
        );
      }
      return out.join(" || ");
    });
  await step(
    "ppt pane",
    async () =>
      (await pp.evaluate(
        () => (document.body.innerText.match(/v\d+\.\d+\.\d+/) || [""])[0],
      )) +
      " " +
      (await pp.evaluate(() =>
        document.body.innerText.replace(/\s+/g, " ").slice(0, 60),
      )),
  );
  await step("select slide 3", () =>
    prun(async (c) => {
      const s = c.presentation.slides.getItemAt(2);
      s.load("id");
      await c.sync();
      c.presentation.setSelectedSlides([s.id]);
      await c.sync();
      return s.id;
    }),
  );
  await step("inbox", async () => {
    await pp.locator("#tab-inbox").click();
    await pp
      .locator("#refresh-inbox")
      .click({ timeout: 8000 })
      .catch(() => {});
    await ppt.waitForTimeout(7000);
    return JSON.stringify(
      await pp.evaluate(() =>
        Array.from(document.querySelectorAll(".inbox-row")).map((r) =>
          r.innerText.replace(/\s+/g, " ").slice(0, 40),
        ),
      ),
    );
  });
  for (let i = 0; i < 3; i++) {
    await step("insert #" + (i + 1), async () => {
      const b = pp.locator(".inbox-insert").first();
      if (!(await b.count())) return "nothing to insert";
      const t0 = Date.now();
      await b.click({ timeout: 8000 });
      for (let w = 0; w < 40; w++) {
        await ppt.waitForTimeout(3000);
        const t = await ptoast();
        if (/Inserted|picture|shapes|budget|Placed|over/i.test(t))
          return Math.round((Date.now() - t0) / 1000) + "s toast: " + t;
      }
      return (
        Math.round((Date.now() - t0) / 1000) + "s no toast: " + (await ptoast())
      );
    });
  }
  await step("slide 3 after inserts", slide3);
  await ppt.screenshot({ path: out + "proof-inserted.png" });
  await excel.bringToFront();
  await step("change source + push all", async () => {
    await xrun(async (c) => {
      const r = c.workbook.worksheets.getItem("P&L").getRange("C11");
      r.load("values");
      await c.sync();
      r.values = [[Number(r.values[0][0]) + 500]];
      await c.sync();
    });
    await xtab("Links");
    await xp.locator("#push-all").click({ timeout: 15000 });
    await excel.waitForTimeout(30000);
    return await xtoast();
  });
  await ppt.bringToFront();
  await step("update all", async () => {
    await pp.locator("#tab-links").click();
    await pp.locator("#refresh-links").click({ timeout: 8000 });
    await ppt.waitForTimeout(9000);
    const before = await pp.evaluate(
      () => (document.body.innerText.match(/Update available/g) || []).length,
    );
    await pp.locator("#update-all").click({ timeout: 8000 });
    for (let w = 0; w < 40; w++) {
      await ppt.waitForTimeout(4000);
      const t = await ptoast();
      if (/updated|Updated|failed|Up to date/i.test(t))
        return before + " rows were 'Update available' -> " + t;
    }
    return "no toast after 160s: " + (await ptoast());
  });
  await step("slide 3 after update", slide3);
  await ppt.screenshot({ path: out + "proof-updated.png" });
  return log;
};
