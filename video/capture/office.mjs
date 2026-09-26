// office.mjs: the moves every capture makes in Excel and PowerPoint for the web: open a
// document at the capture size with the add-in registered, reach the pane frame, run
// Office.js inside it, press a real pane button and wait for its toast, and take one
// screenshot with the Office header (account, file path) cropped away. Invariant: every
// rectangle is capture CSS px of the cropped image, so the composition never sees the header.
import { writeFileSync } from "node:fs";

export const VIEW = { w: 1440, h: 948, dpr: 2 };
export const HEADER = 48; // the Office header band, cropped from every shot
export const PANE = { x: /modelis\/taskpane/, p: /modelis\/pptpane/ };

const wait = (page, ms) => page.waitForTimeout(ms);

// A frame by its real location (cross-origin frames report an empty url() over CDP).
export async function frameByUrl(page, re, tries = 20) {
  for (let i = 0; i < tries; i++) {
    for (const f of page.frames()) {
      const href = await f.evaluate(() => location.href).catch(() => "");
      if (re.test(href)) return f;
    }
    await wait(page, 500);
  }
  throw new Error(`frame ${re} not found`);
}

// Opens a document at the capture size, answers the add-in registration dialogs and waits
// until the ribbon carries the pls,fix tab. The developer opt-in only counts on the next load
// of the same URL (scripts/rig/README.md), so a document gets up to three loads.
export async function openDoc(ctx, url, label) {
  const page = await ctx.newPage();
  const cdp = await ctx.newCDPSession(page);
  page.__cdp = cdp; // shoot() captures through it, at the emulated device scale
  await cdp.send("Emulation.setDeviceMetricsOverride", { width: VIEW.w, height: VIEW.h, deviceScaleFactor: VIEW.dpr, mobile: false });
  await cdp.send("Browser.grantPermissions", { permissions: LOCAL });
  // a page that thinks it is in the background throttles its timers and the pane with them
  await cdp.send("Emulation.setFocusEmulationEnabled", { enabled: true });
  await page.bringToFront();
  for (let load = 1; load <= 3; load++) {
    if (load === 1) await page.goto(url, { waitUntil: "domcontentloaded", timeout: 120000 });
    else await page.reload({ waitUntil: "domcontentloaded", timeout: 120000 });
    for (let i = 0; i < 20; i++) {
      await wait(page, 3000);
      await answerDialogs(page);
      if (await hasTab(page, /^pls,fix$/)) {
        await wait(page, 2500);
        await dismiss(page);
        return page;
      }
    }
    console.log(`capture: ${label}: no pls,fix tab after load ${load}, loading again`);
    await grantFrames(page, cdp);
  }
  throw new Error(`open ${label}: the pls,fix ribbon tab never appeared`);
}

// Chrome asks before a public page reaches 127.0.0.1 (the manifest server); a headless page
// cannot answer, so the grant goes to every origin the document's frames run on. Chrome 153
// names three permissions for it; 127.0.0.1 itself is `loopbackNetwork`.
const LOCAL = ["localNetworkAccess", "localNetwork", "loopbackNetwork"];
async function grantFrames(page, cdp) {
  const origins = new Set();
  for (const f of page.frames()) {
    const o = await f.evaluate(() => location.origin).catch(() => "");
    if (/^https:/.test(o)) origins.add(o);
  }
  for (const origin of origins) {
    await cdp.send("Browser.grantPermissions", { origin, permissions: LOCAL }).catch(() => {});
  }
}

async function answerDialogs(page) {
  for (const f of page.frames()) {
    const t = await f.evaluate(() => (document.body ? document.body.innerText : "")).catch(() => "");
    if (/Registering Developer Add-in Manifest/i.test(t)) await clickRole(f, "button", /^Yes$/);
    else if (/Enable Developer Mode/i.test(t)) {
      const cb = f.locator("#optInCheckbox, #WACDialogOptInCheckbox-input").first();
      if ((await cb.count()) && !(await cb.isChecked())) await f.getByText(/Enable Developer Mode now/i).first().click({ force: true }).catch(() => {});
      await clickRole(f, "button", /^OK$/);
    } else if (/Cannot access manifest url/i.test(t)) await clickRole(f, "button", /^OK$/);
  }
}

async function clickRole(frame, role, name) {
  const l = frame.getByRole(role, { name }).first();
  if (await l.count().catch(() => 0)) {
    await l.click({ timeout: 5000 }).catch(() => {});
    return true;
  }
  return false;
}

export async function hasTab(page, name) {
  for (const f of page.frames()) {
    if (await f.getByRole("tab", { name }).count().catch(() => 0)) return true;
  }
  return false;
}

// Clicks a ribbon control (tab or button) by its accessible name, in whichever frame holds it.
export async function ribbon(page, role, name) {
  for (const f of page.frames()) {
    if (await clickRole(f, role, name)) {
      await wait(page, 900);
      return;
    }
  }
  throw new Error(`ribbon ${role} ${name} not found`);
}

// The page box of a ribbon control, for the pointer (null when the ribbon hides it).
export async function ribbonBox(page, role, name) {
  for (const f of page.frames()) {
    const l = f.getByRole(role, { name }).first();
    if (await l.count().catch(() => 0)) return l.boundingBox();
  }
  return null;
}

// Closes Office's own pop-ups (feedback asks, teaching bubbles) that would land in a shot.
export async function dismiss(page) {
  for (const f of page.frames()) {
    await f.evaluate(() => {
      const closers = [...document.querySelectorAll("button[aria-label], [role=button][aria-label]")]
        .filter((b) => /^(close|dismiss)$/i.test(b.getAttribute("aria-label") || ""));
      for (const b of closers) {
        const box = b.closest("[role=dialog], [role=alertdialog], [class*=eachingBubble], [class*=allout], [class*=eedback]");
        if (box && !/pls,fix/i.test(box.innerText || "")) b.click();
      }
    }).catch(() => {});
  }
  await wait(page, 400);
}

// Opens the pane from the ribbon when it is not showing, and answers its frame.
export async function pane(page, kind, button) {
  const re = PANE[kind];
  const visible = async () => {
    for (const f of page.frames()) {
      const href = await f.evaluate(() => location.href).catch(() => "");
      if (!re.test(href)) continue;
      const box = await (await f.frameElement()).boundingBox().catch(() => null);
      if (box && box.width > 50) return f;
    }
    return null;
  };
  let f = await visible();
  if (!f) {
    await ribbon(page, "tab", /^pls,fix$/);
    await ribbon(page, "button", button);
    for (let i = 0; i < 30 && !f; i++) {
      await wait(page, 1000);
      f = await visible();
    }
  }
  if (!f) throw new Error(`pane ${kind} did not open`);
  await f.waitForFunction(() => document.querySelector("#tab-bar"), null, { timeout: 60000 });
  return f;
}

// Freezes the pane's CSS transitions (a shot never lands mid-fade) and logs every toast the
// pane shows, so a press can wait for its own receipt however briefly it stays up.
export async function prime(frame) {
  await frame.evaluate(() => {
    if (window.__toastLog) return;
    const style = document.createElement("style");
    style.textContent = "*, *::before, *::after { transition: none !important; animation: none !important; }";
    document.head.append(style);
    window.__toastLog = [];
    const t = document.getElementById("toast");
    new MutationObserver(() => {
      if (/visible/.test(t.className)) {
        const text = t.innerText.replace(/\s+/g, " ").trim();
        const last = window.__toastLog[window.__toastLog.length - 1];
        if (!last || last.text !== text || last.cls !== t.className) window.__toastLog.push({ text, cls: t.className });
      }
    }).observe(t, { attributes: true, childList: true, subtree: true, characterData: true });
  });
}

// Presses a pane control and answers the toast it raised ({text, cls}) once the pane is idle.
export async function press(page, frame, selector, { wait: waitMs = 60000 } = {}) {
  await prime(frame);
  const n0 = await frame.evaluate(() => window.__toastLog.length);
  await frame.evaluate((sel) => {
    const el = document.querySelector(sel);
    if (!el) throw new Error(`press: ${sel} not found`);
    el.scrollIntoView({ block: "center" });
    el.click();
  }, selector);
  const t0 = Date.now();
  while (Date.now() - t0 < waitMs) {
    await wait(page, 300);
    const r = await frame.evaluate((n0) => {
      const busy = [...document.querySelectorAll("#tab-bar button")].some((b) => b.disabled);
      const log = window.__toastLog;
      return { busy, last: log.length > n0 ? log[log.length - 1] : null };
    }, n0);
    if (r.last && !r.busy) return r.last;
  }
  throw new Error(`press ${selector}: no toast within ${waitMs / 1000} s`);
}

// Shows a toast the pane raised a moment ago again (its own text and kind), for the shot.
export async function pin(frame, toast) {
  await frame.evaluate((t) => {
    const el = document.getElementById("toast");
    if (el.innerText.replace(/\s+/g, " ").trim() !== t.text) throw new Error(`pin: toast now reads ${el.innerText}`);
    el.className = t.cls;
  }, toast);
}

// Hides the toast (its own timer would a moment later), so a "before" shot is clean.
export async function unpin(frame) {
  await prime(frame);
  await frame.evaluate(() => { document.getElementById("toast").className = "toast"; });
}

export async function toastText(frame) {
  return frame.evaluate(() => {
    const t = document.querySelector("#toast");
    return t && /visible/.test(t.className) ? t.innerText.replace(/\s+/g, " ").trim() : "";
  }).catch(() => "");
}

// Opens one of the pane's own tabs.
export async function tab(frame, id) {
  await frame.evaluate((id) => document.getElementById(id).click(), id);
  await frame.page().waitForTimeout(1500);
}

// Sets a pane <select> or <input> the way a person does (value, then input and change events).
export async function setField(frame, id, value) {
  await frame.evaluate(({ id, value }) => {
    const el = document.getElementById(id);
    if (!el) throw new Error(`field ${id} not found`);
    el.value = String(value);
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  }, { id, value });
  await frame.page().waitForTimeout(600);
}

// Scrolls the pane so the element sits `top` px below the pane's top edge.
export async function scrollPane(frame, selector, top = 12) {
  await frame.evaluate(({ sel, top }) => {
    const el = document.querySelector(sel);
    const shell = document.querySelector(".app-shell");
    if (!el || !shell) throw new Error(`scroll: ${sel} not found`);
    shell.scrollTop += el.getBoundingClientRect().top - shell.getBoundingClientRect().top - top;
  }, { sel: selector, top });
  await frame.page().waitForTimeout(500);
}

// One screenshot, the Office header cropped away: 1440 x 900 CSS px at 2x (2880 x 1800).
export async function shoot(page, path) {
  await dismiss(page);
  await wait(page, 700);
  const clip = { x: 0, y: HEADER, width: VIEW.w, height: VIEW.h - HEADER, scale: 1 };
  const { data } = await page.__cdp.send("Page.captureScreenshot", { format: "png", clip, captureBeyondViewport: false });
  writeFileSync(path, Buffer.from(data, "base64"));
}

// A page box as a capture rect [x, y, w, h] (header removed), rounded to 0.5 px.
export function rect(box) {
  const r = (v) => Math.round(v * 2) / 2;
  return [r(box.x), r(box.y - HEADER), r(box.width), r(box.height)];
}

// Rects of pane elements (the first visible match, or all visible matches when `all`), clipped
// to what the pane shows: a long list runs on below the fold, and only its visible part is real.
export async function paneRects(frame, specs) {
  const fb = await (await frame.frameElement()).boundingBox();
  const view = { x0: fb.x, y0: Math.max(fb.y, HEADER), x1: Math.min(fb.x + fb.width, VIEW.w), y1: Math.min(fb.y + fb.height, VIEW.h) };
  const out = {};
  for (const [key, sel, all] of specs) {
    const boxes = [];
    for (const l of await frame.locator(sel).all()) {
      const b = await l.boundingBox();
      if (!b) continue;
      const x0 = Math.max(b.x, view.x0), y0 = Math.max(b.y, view.y0);
      const x1 = Math.min(b.x + b.width, view.x1), y1 = Math.min(b.y + b.height, view.y1);
      if (x1 - x0 >= 4 && y1 - y0 >= 4) boxes.push(rect({ x: x0, y: y0, width: x1 - x0, height: y1 - y0 }));
      if (boxes.length && !all) break;
    }
    if (!boxes.length) throw new Error(`rect ${key}: ${sel} matched nothing visible`);
    out[key] = all ? boxes : boxes[0];
  }
  return out;
}
