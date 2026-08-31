// Helpers every drive.mjs snippet gets as `h`: find the Excel or PowerPoint
// host page, find a frame by its real location (pre-existing cross-origin
// frames report an empty url() over a fresh CDP connection), click by role in
// any frame, list ribbon tabs, and the screenshot folder.
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));

// Screenshots and other run output: git-ignored, next to the UX gate's own.
export function shotsDir() {
  return join(here, "..", "..", ".superpowers", "rig");
}

export async function frameByUrl(page, re, tries = 6) {
  for (let i = 0; i < tries; i++) {
    for (const f of page.frames()) {
      const href = await f.evaluate(() => location.href).catch(() => "");
      if (re.test(href)) return f;
    }
    await page.waitForTimeout(500);
  }
  return null;
}

// kind: "x" for Excel, "p" for PowerPoint (the SharePoint document URL scheme).
export function hostPage(ctx, kind) {
  return ctx
    .pages()
    .filter(
      (p) =>
        new RegExp("sharepoint\\.com/:" + kind + ":").test(p.url()) &&
        !p.isClosed(),
    )
    .sort((a, b) => b.frames().length - a.frames().length)[0];
}

export async function paneText(page, re, limit = 400) {
  const f = await frameByUrl(page, re);
  if (!f) return null;
  return f.evaluate(
    (n) => document.body.innerText.replace(/\s+/g, " ").slice(0, n),
    limit,
  );
}

export async function clickIn(page, role, name) {
  for (const f of page.frames()) {
    const l = f.getByRole(role, { name }).first();
    if (await l.count().catch(() => 0)) {
      await l.click();
      return true;
    }
  }
  return false;
}

export async function ribbonTabs(page) {
  const tabs = [];
  for (const f of page.frames()) {
    const t = await f
      .getByRole("tab")
      .allTextContents()
      .catch(() => []);
    tabs.push(...t.map((s) => s.trim()).filter(Boolean));
  }
  return [...new Set(tabs)];
}
