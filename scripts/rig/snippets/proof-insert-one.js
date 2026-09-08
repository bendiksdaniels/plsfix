// 08.09 links proof: insert the first inbox item on slide SLIDE (0-based, added when
// missing), INSERTS times, polling busy/toast/shape count every 10 s so a jammed batch
// shows as "JAMMED" instead of a silent wait. Heal a jam by reloading the pane frame.
async (page, ctx, pages, shot, h) => {
  const ppt = h.hostPage(ctx, "p");
  const log = [];
  await ppt.bringToFront();
  await ppt.waitForTimeout(800);
  const pp = await h.frameByUrl(ppt, /modelis\/pptpane/, 3);
  if (!pp) return ["no pane"];
  const prun = (fn, arg, ms = 60000) =>
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
  const ptoast = () =>
    pp.evaluate(() =>
      (document.querySelector("#toast")?.textContent || "")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 240),
    );
  const busy = () =>
    pp.evaluate(() => !!document.querySelector("#tab-inbox")?.disabled);
  const count = () =>
    prun(
      async (c) => {
        const s = c.presentation.slides.getItemAt(
          Number(process.env.SLIDE || 0),
        ).shapes;
        s.load("items/id");
        await c.sync();
        return s.items.length;
      },
      null,
      20000,
    ).catch(() => "?");
  const slideIndex = Number(process.env.SLIDE || 0);
  log.push(
    "select slide " +
      (slideIndex + 1) +
      ": " +
      (await prun(
        async (c, i) => {
          const slides = c.presentation.slides;
          slides.load("items/id");
          await c.sync();
          while (slides.items.length <= i) {
            slides.add();
            await c.sync();
            slides.load("items/id");
            await c.sync();
          }
          const s = slides.getItemAt(i);
          s.load("id");
          await c.sync();
          c.presentation.setSelectedSlides([s.id]);
          await c.sync();
          return s.id;
        },
        slideIndex,
        40000,
      )),
  );
  await pp.evaluate(() => document.querySelector("#tab-inbox")?.click());
  await ppt.waitForTimeout(800);
  const n = Number(process.env.INSERTS || 1);
  for (let i = 0; i < n; i++) {
    const b = pp.locator(".inbox-insert").first();
    if (!(await b.count())) {
      log.push("inbox empty");
      break;
    }
    const label = await b.evaluate((e) =>
      (e.closest("li, tr, div")?.innerText || "")
        .replace(/\s+/g, " ")
        .slice(0, 28),
    );
    const before = await ptoast();
    const t0 = Date.now();
    const c0 = await count();
    await b.evaluate((e) => e.click());
    let result = "no result";
    for (let w = 0; w < 24; w++) {
      await ppt.waitForTimeout(10000);
      const t = await ptoast();
      const bz = await busy();
      const cnt = await count();
      log.push(
        "  " +
          label +
          " +" +
          Math.round((Date.now() - t0) / 1000) +
          "s busy=" +
          bz +
          " shapes=" +
          cnt +
          (t !== before ? " toast=" + t : ""),
      );
      if (!bz && t !== before) {
        result =
          label + " -> " + Math.round((Date.now() - t0) / 1000) + "s: " + t;
        break;
      }
    }
    log.push(result + " (shapes " + c0 + " -> " + (await count()) + ")");
    if (result === "no result") {
      log.push("JAMMED: stopping");
      break;
    }
  }
  await ppt.screenshot({
    path: h.shotsDir() + "/proof-c-slide" + (slideIndex + 1) + ".png",
  });
  return log;
};
