// run.mjs: the capture, `node video/capture/run.mjs <out-dir>` (plsfix-video capture runs it):
// connects to the rig's scratch Chrome on CDP 9222, uploads the fresh demo workbook and deck,
// opens each in Office for the web with pls,fix registered, runs the Excel and PowerPoint shot
// lists and writes <out-dir>/capture.json: every shot's file and rects, the pane version read
// off the pane footer, and each uploaded file's sha256. The site comes from PLSFIX_RIG_SITE.
import { createHash } from "node:crypto";
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { loadPlaywright } from "../../scripts/ux/browser.mjs";
import { upload } from "./upload.mjs";
import { openDoc, pane, shoot } from "./office.mjs";
import { excelShots } from "./excel-shots.mjs";
import { linkShots } from "./ppt-shots.mjs";

const REPO = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
const WORKBOOK = join(REPO, "demo/out/pls,fix Demo Model.xlsx");
const DECK = join(REPO, "demo/deck/pls,fix Demo Deck.pptx");
const CDP = process.env.RIG_CDP ?? "http://127.0.0.1:9222";

const out = process.argv[2];
const site = process.env.PLSFIX_RIG_SITE;
if (!out || !site) {
  console.error("usage: PLSFIX_RIG_SITE=<onedrive site> node video/capture/run.mjs <out-dir>");
  process.exit(2);
}
mkdirSync(out, { recursive: true });
const sha = (p) => createHash("sha256").update(readFileSync(p)).digest("hex");

const { chromium } = loadPlaywright();
const browser = await chromium.connectOverCDP(CDP);
const ctx = browser.contexts()[0];
const shots = {};
const snapper = (page) => async (name, rects) => {
  await shoot(page, join(out, `${name}.png`));
  shots[name] = { file: `${name}.png`, rects };
  console.log(`capture: ${name} (${Object.keys(rects).length} rects)`);
};

try {
  // Earlier Office tabs of the rig would keep their own add-in runtimes busy: close them.
  for (const p of ctx.pages()) if (/sharepoint\.com/.test(p.url())) await p.close();
  const [wb, deck] = await upload(ctx, site, [WORKBOOK, DECK]);
  console.log(`capture: uploaded ${wb.name} (${wb.bytes} B) and ${deck.name} (${deck.bytes} B)`);
  const xl = await openDoc(ctx, wb.url, "workbook");
  const xf = await pane(xl, "x", /^Model Tools$/);
  const version = await xf.evaluate(() => document.getElementById("app-version").textContent.trim());
  await excelShots(xl, xf, snapper(xl));
  const pp = await openDoc(ctx, deck.url, "deck");
  await linkShots({ xl, xf, pp, snap: snapper, openPane: pane });
  const record = {
    source: "npm run demo:build",
    host: "Office for the web",
    pane_version: version,
    date: new Date().toISOString().slice(0, 10),
    workbook: { name: wb.name, sha256: sha(WORKBOOK) },
    deck: { name: deck.name, sha256: sha(DECK) },
    shots,
  };
  writeFileSync(join(out, "capture.json"), JSON.stringify(record, null, 1));
  console.log(`capture: ${Object.keys(shots).length} shots, pane ${version}`);
} finally {
  await browser.close(); // disconnects; the scratch Chrome keeps running
}
