#!/usr/bin/env node
// The 23.09 clipboard spike made permanent: proof, in real headless
// Chromium (never jsdom, which has no Clipboard API), that a bundle written
// with beginClipboardWrite and with copyBundleNow (src/ui/clipboard-links.ts)
// both round-trip through the real OS/browser clipboard and read back with
// bundleFromPaste on a second page of the same origin - the two halves of
// docs/superpowers/specs/2026-09-23-local-links-design.md s.2's carrier.
// Both pages load taskpane.html only as a Vite-served host: src/link/
// bundle.ts and src/ui/clipboard-links.ts are imported for real through
// "/src/..." (the way scripts/ux/panes.mjs imports pane modules), never
// copied or reimplemented here.
//
// FAILs unless: beginClipboardWrite lands a >= 5 MB bundle byte for byte;
// copyBundleNow does too, independently; the plain-text flavor is
// BUNDLE_SENTENCE both times; a real Ctrl+V (Cmd+V on a Mac) is what
// delivers the paste event, not a synthetic one.
//
// Usage: npm run ux:clipboard [-- --port 3212]
// Own port (per the brief): 3212, never the ux:sweep/ux:check ports.

import { launchBrowser } from "./browser.mjs";
import { ensureServer } from "./dev-server.mjs";

const OWN_PORT = 3212;
// The design's own proof size (spec s.2: "both write paths carry 5 MB in
// the HTML flavor").
const BUNDLE_BLOB_BYTES = 5 * 1024 * 1024;
const PASTE_KEYS = process.platform === "darwin" ? "Meta+v" : "Control+v";

function flag(argv, name, fallback) {
  const at = argv.indexOf(`--${name}`);
  return at === -1 ? fallback : argv[at + 1];
}

function parseArgs(argv) {
  const port = Number(flag(argv, "port", OWN_PORT));
  if (!Number.isInteger(port) || port <= 0) {
    throw new Error(`--port needs an integer, got ${String(port)}`);
  }
  return { port };
}

// A real page from the dev server, so "/src/..." imports resolve through
// Vite's own module graph - never navigated for its own UI, only as a
// same-origin host both the write and the read run against.
async function openHost(context, baseUrl) {
  const page = await context.newPage();
  await page.goto(`${baseUrl}/taskpane.html`, {
    waitUntil: "domcontentloaded",
  });
  return page;
}

// One >= 5 MB bundle, built and hashed entirely in the page (crypto.
// getRandomValues caps a single call at 65536 bytes, filled in chunks; the
// blob never crosses back to Node - only its short digest does, so a
// multi-megabyte string is never round-tripped through evaluate()'s own
// JSON marshalling).
async function buildBundle(page, id, targetBytes) {
  return page.evaluate(
    async ([linkId, size]) => {
      const { encodeBundle } = await import("/src/link/bundle.ts");
      const blob = new Uint8Array(size);
      const CHUNK = 65536;
      for (let offset = 0; offset < size; offset += CHUNK) {
        crypto.getRandomValues(blob.subarray(offset, offset + CHUNK));
      }
      const digestBytes = await crypto.subtle.digest("SHA-256", blob);
      const digest = [...new Uint8Array(digestBytes)]
        .map((byte) => byte.toString(16).padStart(2, "0"))
        .join("");
      const bundle = {
        links: [{ id: linkId, rev: 2 ** 40 + 1, sentAt: Date.now(), blob }],
        inbox: [],
      };
      const json = encodeBundle(bundle);
      return { json, digest, length: json.length };
    },
    [id, targetBytes],
  );
}

// beginClipboardWrite() must start synchronously inside a real press (the
// code's own invariant, WebKit-driven); a page.evaluate() call carries no
// user activation in Chromium, so the button exists to give it one for
// real, exactly the way the Excel pane's own click handler does.
async function installWriteButtons(page) {
  await page.evaluate(async () => {
    const { beginClipboardWrite, copyBundleNow } =
      await import("/src/ui/clipboard-links.ts");
    window.__probe = { beginClipboardWrite, copyBundleNow, pending: null };
    const begin = document.createElement("button");
    begin.id = "probe-begin-write";
    begin.onclick = () => {
      window.__probe.pending = window.__probe.beginClipboardWrite();
    };
    const copyNow = document.createElement("button");
    copyNow.id = "probe-copy-now";
    copyNow.onclick = () => {
      window.__probe.copyNowResult = window.__probe.copyBundleNow(
        window.__probe.copyNowJson,
      );
    };
    document.body.append(begin, copyNow);
  });
}

// Round 1: the one-click async write. beginClipboardWrite() runs inside the
// real click above; resolving it with the bundle (the content is allowed to
// arrive after the press, per the design) and awaiting `done` happens after,
// same as src/pane/links-transport.ts's own copy().
async function writeWithBeginClipboardWrite(page, json) {
  await page.click("#probe-begin-write");
  return page.evaluate(async (text) => {
    window.__probe.pending.resolve(text);
    return window.__probe.pending.done;
  }, json);
}

// Round 2: the synchronous execCommand("copy") fallback, driven by its own
// real click so it never depends on round 1's activation.
async function writeWithCopyBundleNow(page, json) {
  await page.evaluate((text) => {
    window.__probe.copyNowJson = text;
  }, json);
  await page.click("#probe-copy-now");
  return page.evaluate(() => window.__probe.copyNowResult);
}

// A real Ctrl+V (Cmd+V on a Mac) into a focused textarea, read back with the
// real bundleFromPaste/readPastedBundle - never a synthetic paste Event,
// which would only prove the reader, not the carrier a real user presses.
async function installPasteBox(page) {
  await page.evaluate(() => {
    let box = document.getElementById("probe-paste-box");
    if (!box) {
      box = document.createElement("textarea");
      box.id = "probe-paste-box";
      document.body.append(box);
    }
    box.value = "";
    window.__pasted = new Promise((resolve) => {
      box.addEventListener(
        "paste",
        (event) => {
          resolve({
            html: event.clipboardData?.getData("text/html") ?? "",
            plain: event.clipboardData?.getData("text/plain") ?? "",
          });
        },
        { once: true },
      );
    });
    box.focus();
  });
}

// The plain-text expectation is read out of the real module too (never a
// copy of the sentence kept here to drift from src/link/bundle.ts).
async function readBack(page, expectedDigest) {
  await installPasteBox(page);
  await page.keyboard.press(PASTE_KEYS);
  return page.evaluate(async (digest) => {
    const { BUNDLE_SENTENCE, readPastedBundle } =
      await import("/src/link/bundle.ts");
    const captured = await window.__pasted;
    const read = readPastedBundle(captured.html, captured.plain);
    if (!read.ok) {
      return { ok: false, reason: `decode failed: ${read.reason}` };
    }
    const link = read.bundle.links[0];
    if (link === undefined) {
      return { ok: false, reason: "decoded bundle carried no links" };
    }
    const digestBytes = await crypto.subtle.digest("SHA-256", link.blob);
    const gotDigest = [...new Uint8Array(digestBytes)]
      .map((byte) => byte.toString(16).padStart(2, "0"))
      .join("");
    if (gotDigest !== digest) {
      return { ok: false, reason: "blob digest changed in the round trip" };
    }
    if (captured.plain !== BUNDLE_SENTENCE) {
      return {
        ok: false,
        reason: `plain text was "${captured.plain}", want "${BUNDLE_SENTENCE}"`,
      };
    }
    return { ok: true };
  }, expectedDigest);
}

async function runRound(label, write, writePage, readPage, id, failures) {
  const built = await buildBundle(writePage, id, BUNDLE_BLOB_BYTES);
  process.stdout.write(
    `${label}: built a ${String(built.length)}-char bundle (blob ${String(BUNDLE_BLOB_BYTES)} bytes)\n`,
  );
  const landed = await write(writePage, built.json);
  if (landed !== true) {
    failures.push(
      `${label}: the write did not report landing (got ${String(landed)})`,
    );
    return;
  }
  const result = await readBack(readPage, built.digest);
  if (!result.ok) {
    failures.push(`${label}: ${result.reason}`);
    return;
  }
  process.stdout.write(
    `${label}: round-tripped ${String(built.length)} chars, digest matched\n`,
  );
}

async function run() {
  const options = parseArgs(process.argv.slice(2));
  const server = await ensureServer(options.port);
  const browser = await launchBrowser();
  const failures = [];
  try {
    const context = await browser.newContext({ ignoreHTTPSErrors: true });
    await context.grantPermissions(["clipboard-read", "clipboard-write"], {
      origin: new URL(server.url).origin,
    });
    const writePage = await openHost(context, server.url);
    const readPage = await openHost(context, server.url);
    await installWriteButtons(writePage);

    await runRound(
      "beginClipboardWrite",
      writeWithBeginClipboardWrite,
      writePage,
      readPage,
      "a".repeat(32),
      failures,
    );
    await runRound(
      "copyBundleNow",
      writeWithCopyBundleNow,
      writePage,
      readPage,
      "b".repeat(32),
      failures,
    );

    await context.close();
  } finally {
    await browser.close();
    server.stop();
  }

  if (failures.length === 0) {
    process.stdout.write("PASS - both write paths round-tripped 5 MB\n");
    process.exitCode = 0;
    return;
  }
  for (const failure of failures) process.stdout.write(`${failure}\n`);
  process.stdout.write(`FAIL - ${String(failures.length)} defect(s)\n`);
  process.exitCode = 1;
}

run().catch((error) => {
  process.stderr.write(`${String(error?.stack ?? error)}\n`);
  process.exitCode = 1;
});
