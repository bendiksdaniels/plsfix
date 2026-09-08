// 08.09 links proof, update half: bump P&L!C4 and Proof!B2 in Excel, Push all, then
// Update all in PowerPoint and read every link's rev back from its tag.
async (page, ctx, pages, shot, h) => {
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
            .slice(0, 220),
      );
    }
  };
  const xp = await h.frameByUrl(excel, /modelis\/taskpane/, 2);
  if (!xp) return ["no excel pane"];
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
  const xclick = (sel) =>
    xp.evaluate((sel) => {
      const el = document.querySelector(sel);
      if (!el) throw new Error("no " + sel);
      el.click();
    }, sel);
  await excel.bringToFront();
  await excel.waitForTimeout(800);
  await step("change sources", () =>
    xrun(async (c) => {
      const r = c.workbook.worksheets.getItem("P&L").getRange("C4");
      r.load("values");
      await c.sync();
      r.values = [[Number(r.values[0][0]) + 500]];
      const p = c.workbook.worksheets.getItem("Proof").getRange("B2");
      p.load("values");
      await c.sync();
      p.values = [[Number(p.values[0][0]) + 40]];
      await c.sync();
      return "P&L!C4 +500, Proof!B2 +40";
    }),
  );
  await step("push all", async () => {
    await xp
      .getByRole("tab", { name: /^Links$/ })
      .first()
      .click({ timeout: 20000, force: true })
      .catch(() => {});
    await excel.waitForTimeout(1200);
    const before = await xtoast();
    await xclick("#push-all");
    for (let w = 0; w < 30; w++) {
      await excel.waitForTimeout(3000);
      const t = await xtoast();
      if (t !== before && /pushed|missing|failed/i.test(t)) return t;
    }
    return "no push toast: " + (await xtoast());
  });
  await ppt.bringToFront();
  await ppt.waitForTimeout(800);
  const pp = await h.frameByUrl(ppt, /modelis\/pptpane/, 3);
  if (!pp) return log.concat(["no ppt pane"]);
  const pclick = (sel) =>
    pp.evaluate((sel) => {
      const el = document.querySelector(sel);
      if (!el) throw new Error("no " + sel);
      el.click();
    }, sel);
  const ptoast = () =>
    pp.evaluate(() =>
      (document.querySelector("#toast")?.textContent || "")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 240),
    );
  const prun = (fn, arg, ms = 120000) =>
    pp.evaluate(
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
  await step("links list before", async () => {
    await pclick("#tab-links");
    await ppt.waitForTimeout(1000);
    await pclick("#refresh-links");
    await ppt.waitForTimeout(12000);
    return (
      (await pp.evaluate(() =>
        Array.from(
          document.querySelectorAll(
            "#view-links tbody tr, #view-links .wl-row, #view-links li",
          ),
        )
          .map((r) => r.innerText.replace(/\s+/g, " ").slice(0, 70))
          .join(" || "),
      )) ||
      (await pp.evaluate(() =>
        (document.querySelector("#view-links")?.innerText || "")
          .replace(/\s+/g, " ")
          .slice(200, 600),
      ))
    );
  });
  await step("update all", async () => {
    const before = await ptoast();
    const t0 = Date.now();
    await pclick("#update-all");
    for (let w = 0; w < 90; w++) {
      await ppt.waitForTimeout(4000);
      const t = await ptoast();
      if (t !== before && /updated|up to date|failed|missing/i.test(t))
        return Math.round((Date.now() - t0) / 1000) + "s: " + t;
    }
    return "no toast after 360s: " + (await ptoast());
  });
  await step(
    "details",
    async () =>
      await pp.evaluate(() =>
        (document.querySelector("#toast")?.innerText || "")
          .replace(/\s+/g, " ")
          .slice(0, 400),
      ),
  );
  await step("slide 1 after update", () =>
    prun(
      async (c) => {
        const shapes = c.presentation.slides.getItemAt(0).shapes;
        shapes.load(
          "items/id,items/name,items/type,items/left,items/top,items/width,items/height",
        );
        await c.sync();
        const out = [];
        for (const s of shapes.items) {
          if (s.type === "Placeholder") continue;
          const tags = s.tags;
          tags.load("items/key,items/value");
          let kids = null;
          if (s.type === "Group") {
            kids = s.group.shapes;
            kids.load("items/name,items/type");
          }
          await c.sync();
          const link = tags.items.find((t) => t.key === "PLSFIX_LINK");
          let rev = "?";
          try {
            rev = JSON.parse(link.value).rev;
          } catch {}
          out.push(
            s.type +
              " '" +
              s.name.slice(0, 40) +
              "' " +
              [s.left, s.top, s.width, s.height].map(Math.round).join(",") +
              " rev " +
              rev +
              (kids ? " kids " + kids.items.length : ""),
          );
        }
        return out.join(" || ") || "(none)";
      },
      null,
      90000,
    ),
  );
  await ppt.screenshot({ path: h.shotsDir() + "/proof-b2-ppt.png" });
  return log;
};
