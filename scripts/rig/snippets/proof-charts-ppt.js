// PowerPoint half of the Task 8 proof: heal the host if jammed, select slide 3, insert the
// three chart exports from the inbox, read back groups vs pictures, then Update all after a
// source change pushed from Excel.
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
  await ppt.bringToFront();
  await ppt.waitForTimeout(800);
  const findPane = () => h.frameByUrl(ppt, /modelis\/pptpane/, 3);
  const prunOn = (frame, fn, arg, ms = 90000) =>
    frame.evaluate(
      async ([src, arg, ms]) => {
        const f = new Function("return (" + src + ")")();
        return Promise.race([
          PowerPoint.run((c) => f(c, arg)),
          new Promise((_, rej) =>
            setTimeout(() => rej(new Error("timeout " + ms)), ms),
          ),
        ]);
      },
      [fn.toString(), arg, ms],
    );
  let pp = await findPane();
  const healthy = async () => {
    if (!pp) return false;
    try {
      await prunOn(
        pp,
        async (c) => {
          const s = c.presentation.slides.getItemAt(2).shapes;
          s.load("items/id");
          await c.sync();
          return s.items.length;
        },
        null,
        20000,
      );
      return true;
    } catch {
      return false;
    }
  };
  if (!(await healthy())) {
    log.push("host jammed or pane missing: reloading the deck");
    const base = ppt
      .url()
      .replace(
        /&wdaddindevserverport=[^&]*|&wdaddinmanifestfile=[^&]*|&wdaddinmanifestguid=[^&]*/g,
        "",
      );
    await ppt.goto(base, { waitUntil: "domcontentloaded" });
    await ppt.waitForTimeout(20000);
    await ppt.goto(
      base +
        "&wdaddindevserverport=3001&wdaddinmanifestfile=manifest.prod.xml&wdaddinmanifestguid=FF1B34D8-DD7D-4B39-8FA9-6248CA09DB6E",
      { waitUntil: "domcontentloaded" },
    );
    for (let i = 0; i < 20; i++) {
      await ppt.waitForTimeout(3000);
      for (const f of ppt.frames()) {
        const t = await f
          .evaluate(() => (document.body ? document.body.innerText : ""))
          .catch(() => "");
        if (/Registering Developer Add-in Manifest/i.test(t)) {
          const b = f.getByRole("button", { name: /^Yes$/ }).first();
          if (await b.count().catch(() => 0)) {
            await b.click();
            log.push("registered");
          }
        }
      }
      if ((await h.ribbonTabs(ppt)).some((t) => /pls,fix/.test(t))) break;
    }
    await ppt.waitForTimeout(6000);
    pp = await findPane();
    if (!pp) {
      log.push("tab " + (await h.clickIn(ppt, "tab", /^pls,fix$/)));
      await ppt.waitForTimeout(2500);
      log.push("links " + (await h.clickIn(ppt, "button", /^Links$/)));
      await ppt.waitForTimeout(20000);
      pp = await findPane();
    }
    if (!pp) return log.concat(["no pane after heal"]);
    log.push("healthy after heal: " + (await healthy()));
  }
  const prun = (fn, arg, ms) => prunOn(pp, fn, arg, ms);
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
      return out.join(" || ") || "(empty)";
    });
  await step(
    "ppt pane",
    async () =>
      (await pp.evaluate(
        () => (document.body.innerText.match(/v\d+\.\d+\.\d+/) || [""])[0],
      )) +
      " " +
      (await pp.evaluate(() =>
        document.body.innerText.replace(/\s+/g, " ").slice(0, 50),
      )),
  );
  await step("select slide 3", () =>
    prun(
      async (c) => {
        const s = c.presentation.slides.getItemAt(2);
        s.load("id");
        await c.sync();
        c.presentation.setSelectedSlides([s.id]);
        await c.sync();
        return s.id;
      },
      null,
      40000,
    ),
  );
  await step("inbox", async () => {
    await pp.locator("#tab-inbox").click();
    await ppt.waitForTimeout(1000);
    await pp
      .locator("#refresh-inbox")
      .click({ timeout: 8000 })
      .catch(() => {});
    await ppt.waitForTimeout(9000);
    return (
      JSON.stringify(
        await pp.evaluate(() =>
          Array.from(document.querySelectorAll(".inbox-insert")).map((b) =>
            (b.closest("li, tr, div")?.innerText || "")
              .replace(/\s+/g, " ")
              .slice(0, 40),
          ),
        ),
      ) +
      " | " +
      (await pp.evaluate(() =>
        document.body.innerText
          .replace(/\s+/g, " ")
          .replace(/^.*?Waiting to insert/, "Waiting to insert")
          .slice(0, 160),
      ))
    );
  });
  for (let i = 0; i < 3; i++) {
    await step("insert #" + (i + 1), async () => {
      const b = pp.locator(".inbox-insert").first();
      if (!(await b.count())) return "nothing to insert";
      const t0 = Date.now();
      await b.click({ timeout: 8000 });
      for (let w = 0; w < 50; w++) {
        await ppt.waitForTimeout(3000);
        const t = await ptoast();
        if (t && !/^Inserting/i.test(t))
          return Math.round((Date.now() - t0) / 1000) + "s toast: " + t;
      }
      return (
        Math.round((Date.now() - t0) / 1000) + "s, toast: " + (await ptoast())
      );
    });
  }
  await step("slide 3 after inserts", slide3);
  await ppt.screenshot({ path: out + "proof-inserted.png" });
  const xp = await h.frameByUrl(excel, /modelis\/taskpane/, 2);
  if (xp) {
    await excel.bringToFront();
    await excel.waitForTimeout(800);
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
    await step("change source + push all", async () => {
      await xrun(async (c) => {
        const r = c.workbook.worksheets.getItem("P&L").getRange("C11");
        r.load("values");
        await c.sync();
        r.values = [[Number(r.values[0][0]) + 500]];
        await c.sync();
      });
      const t = xp.getByRole("tab", { name: /^Links$/ }).first();
      await t.click({ timeout: 20000 });
      await excel.waitForTimeout(1200);
      await xp.locator("#push-all").click({ timeout: 15000 });
      await excel.waitForTimeout(35000);
      return await xtoast();
    });
    await ppt.bringToFront();
    await ppt.waitForTimeout(800);
    await step("update all", async () => {
      await pp.locator("#tab-links").click();
      await ppt.waitForTimeout(800);
      await pp.locator("#refresh-links").click({ timeout: 8000 });
      await ppt.waitForTimeout(12000);
      const before = await pp.evaluate(
        () => (document.body.innerText.match(/Update available/g) || []).length,
      );
      await pp.locator("#update-all").click({ timeout: 8000 });
      for (let w = 0; w < 60; w++) {
        await ppt.waitForTimeout(4000);
        const t = await ptoast();
        if (/updated|Updated|failed|up to date|Up to date/i.test(t))
          return before + " rows were 'Update available' -> " + t;
      }
      return "no toast after 240s: " + (await ptoast());
    });
    await step("slide 3 after update", slide3);
    await ppt.screenshot({ path: out + "proof-updated.png" });
  } else log.push("excel pane not found: update step skipped");
  return log;
};
