// Register the add-in in one host (HOST_KIND=x or p) the way 08.09 needed it: navigate with
// the wdaddin* parameters, tick "Enable Developer Mode now", OK, Yes, and reload up to
// twice until the pls,fix tab shows. sideload.js is the single-pass original.
async (page, ctx, pages, shot, h) => {
  const P =
    "&wdaddindevserverport=3001&wdaddinmanifestfile=manifest.prod.xml&wdaddinmanifestguid=FF1B34D8-DD7D-4B39-8FA9-6248CA09DB6E";
  const kind = process.env.HOST_KIND || "x";
  const host = h.hostPage(ctx, kind);
  const log = [];
  const base = host.url().replace(/&wdaddin[^&]*/g, "");
  await host.bringToFront();
  const pass = async (label, secs) => {
    for (let i = 0; i < secs / 3; i++) {
      await host.waitForTimeout(3000);
      for (const f of host.frames()) {
        const t = await f
          .evaluate(() => (document.body ? document.body.innerText : ""))
          .catch(() => "");
        if (/Enable Developer Mode/i.test(t)) {
          const cb = f
            .locator("#optInCheckbox, #WACDialogOptInCheckbox-input")
            .first();
          if (await cb.count().catch(() => 0)) {
            for (
              let k = 0;
              k < 3 && !(await cb.isChecked().catch(() => false));
              k++
            ) {
              const lab = f.getByText(/Enable Developer Mode now/i).first();
              if (await lab.count().catch(() => 0))
                await lab.click({ force: true }).catch(() => {});
              else await cb.click({ force: true }).catch(() => {});
              await host.waitForTimeout(400);
            }
            log.push(
              label + " tick=" + (await cb.isChecked().catch(() => "?")),
            );
          }
          const b = f.getByRole("button", { name: /^OK$/ }).first();
          if (await b.count().catch(() => 0)) {
            await b.click();
            log.push(label + " OK " + i * 3 + "s");
          }
        } else if (/Registering Developer Add-in Manifest/i.test(t)) {
          const b = f.getByRole("button", { name: /^Yes$/ }).first();
          if (await b.count().catch(() => 0)) {
            await b.click();
            log.push(label + " Yes " + i * 3 + "s");
          }
        }
      }
      if (i > 2 && (await h.ribbonTabs(host)).some((t) => /pls,fix/.test(t))) {
        log.push(label + " tab at " + i * 3 + "s");
        return true;
      }
    }
    return false;
  };
  await host.goto(base + P, { waitUntil: "domcontentloaded" });
  let ok = await pass("first", 45);
  if (!ok) {
    await host.goto(base + P, { waitUntil: "domcontentloaded" });
    ok = await pass("reload", 45);
  }
  if (!ok) {
    await host.goto(base + P, { waitUntil: "domcontentloaded" });
    ok = await pass("reload2", 45);
  }
  await host.waitForTimeout(3000);
  const tabs = [...new Set(await h.ribbonTabs(host))];
  await host.screenshot({
    path:
      "/Users/danielsbendiks/plsfix/.superpowers/rig/register-final-" +
      kind +
      ".png",
  });
  return { plsfix: tabs.some((t) => /pls,fix/.test(t)), log };
};
