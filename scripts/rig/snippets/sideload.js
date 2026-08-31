// Register the add-in in an open Excel (HOST_KIND=x) or PowerPoint (HOST_KIND=p) document
// whose URL carries the wdaddin* parameters: grants local-network access over CDP, clicks
// the two registration dialogs, then reports the ribbon tabs. NAV_URL navigates first.
async (page, ctx, pages, shot) => {
  const kind = process.env.HOST_KIND; // "x" excel, "p" powerpoint
  const target = ctx
    .pages()
    .find((p) => new RegExp("sharepoint\\.com/:" + kind + ":").test(p.url()));
  if (!target) return "no page for " + kind;
  const log = [],
    events = [];
  const cdp = await ctx.newCDPSession(target);
  await cdp
    .send("Browser.grantPermissions", { permissions: ["localNetworkAccess"] })
    .catch((e) => log.push("grant: " + e.message.slice(0, 60)));
  target.on("requestfailed", (r) => {
    if (/127\.0\.0\.1|localhost/.test(r.url()))
      events.push(
        "failed " + r.url().slice(0, 60) + " " + (r.failure()?.errorText ?? ""),
      );
  });
  target.on("response", (r) => {
    if (/127\.0\.0\.1|localhost/.test(r.url()))
      events.push("response " + r.status() + " " + r.url().slice(0, 60));
  });
  if (process.env.NAV_URL)
    await target.goto(process.env.NAV_URL, { waitUntil: "domcontentloaded" });
  for (let i = 0; i < 14; i++) {
    await target.waitForTimeout(3000);
    let acted = false;
    for (const f of target.frames()) {
      const t = await f
        .evaluate(() => (document.body ? document.body.innerText : ""))
        .catch(() => "");
      if (/Registering Developer Add-in Manifest/i.test(t)) {
        const b = f.getByRole("button", { name: /^Yes$/ }).first();
        if (await b.count()) {
          await b.click();
          log.push("Yes at " + i * 3 + "s");
          acted = true;
        }
      } else if (/Enable Developer Mode/i.test(t)) {
        const b = f.getByRole("button", { name: /^OK$/ }).first();
        if (await b.count()) {
          await b.click();
          log.push("OK at " + i * 3 + "s");
          acted = true;
        }
      }
      if (acted) break;
    }
  }
  await target.waitForTimeout(12000);
  const tabs = [];
  for (const f of target.frames()) {
    const t = await f
      .getByRole("tab")
      .allTextContents()
      .catch(() => []);
    tabs.push(...t.map((s) => s.trim()).filter(Boolean));
  }
  const notes = [];
  for (const f of target.frames()) {
    const t = await f
      .evaluate(() => (document.body ? document.body.innerText : ""))
      .catch(() => "");
    const m = t.match(
      /[^.\n]*(manifest url|add-in manifest|sideload)[^.\n]*/gi,
    );
    if (m)
      notes.push(
        ...m.slice(0, 3).map((s) => s.replace(/\s+/g, " ").slice(0, 140)),
      );
  }
  await shot("sideload-" + kind);
  return {
    log,
    events: events.slice(0, 8),
    ribbon: [...new Set(tabs)].slice(0, 14),
    plsfix: tabs.some((t) => /pls,fix/i.test(t)),
    notes: [...new Set(notes)].slice(0, 4),
  };
};
