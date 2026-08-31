// Writes public/shortcuts.html from public/shortcuts.json (the hand-kept,
// audited source of every PLSFIX_* action's label and default key combo -
// manifest/spec.ts's ribbon groups only cover a subset and carry no keys) and
// package.json (version): the same printable page the Tools tab's
// shortcut-card button and the first-run card's Shortcut card button both
// open in an Office dialog. --check: exit 1 if the committed file differs
// (CI gate), same shape as build-functions-metadata.ts. Written via temp +
// rename so a reader never sees a half-written file.
import { readFileSync, renameSync, writeFileSync } from "node:fs";
import pkg from "../package.json" with { type: "json" };
import shortcutsSource from "../public/shortcuts.json" with { type: "json" };

interface ShortcutRow {
  label: string;
  combo: string;
}

const escapeHtml = (text: string): string =>
  text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

// Every shortcut in the order shortcuts.json lists it - the same order the
// README's own table follows - paired with the label Office's shortcut
// customization UI shows for that action.
function shortcutRows(): ShortcutRow[] {
  const labels = new Map(
    shortcutsSource.actions.map((action) => [action.id, action.name] as const),
  );
  return shortcutsSource.shortcuts.map((binding) => {
    const label = labels.get(binding.action);
    if (label === undefined) {
      throw new Error(`shortcuts.json: no action named ${binding.action}`);
    }
    return { label, combo: binding.key.default };
  });
}

function renderRow(row: ShortcutRow): string {
  return `      <div class="row"><span>${escapeHtml(row.label)}</span><kbd>${escapeHtml(row.combo)}</kbd></div>`;
}

// Self-contained and print-ready: no scripts, no external requests, so the
// page behaves the same in an Office dialog, a plain browser tab and a print
// preview. Two columns via CSS columns rather than a hand-split array, with
// break-inside: avoid so one shortcut's label and key never split across a
// column or a printed page.
const STYLE = `
      :root { color-scheme: light; }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        padding: 28px 32px;
        color: #14213d;
        background: #fbfcfd;
        font-family: "Aptos", "Segoe UI", system-ui, sans-serif;
      }
      header {
        margin-bottom: 18px;
        border-bottom: 2px solid #14213d;
        padding-bottom: 12px;
      }
      h1 { margin: 0 0 4px; font: 700 20px/1.2 Georgia, serif; }
      .version { margin: 0; color: #5a636e; font-size: 12px; }
      .columns { column-count: 2; column-gap: 32px; }
      .row {
        display: flex;
        align-items: baseline;
        justify-content: space-between;
        gap: 12px;
        padding: 5px 0;
        border-bottom: 1px solid #d8e0e6;
        break-inside: avoid;
      }
      .row span { min-width: 0; font-size: 12px; }
      kbd {
        flex: none;
        padding: 2px 6px;
        border: 1px solid #d8e0e6;
        border-bottom-width: 2px;
        border-radius: 5px;
        color: #11695f;
        background: #f7f9fb;
        font:
          700 11px/1.3 ui-monospace, SFMono-Regular, Menlo, Consolas,
          monospace;
        white-space: nowrap;
      }
      footer { margin-top: 20px; color: #5a636e; font-size: 11px; }
      @media print {
        @page {
          size: A4;
          margin: 14mm;
        }
        body { padding: 0; }
      }
`;

function renderPage(): string {
  const list = shortcutRows().map(renderRow).join("\n");
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>pls,fix keyboard shortcuts</title>
    <style>${STYLE}    </style>
  </head>
  <body>
    <header>
      <h1>pls,fix keyboard shortcuts</h1>
      <p class="version">v${pkg.version}</p>
    </header>
    <div class="columns">
${list}
    </div>
    <footer>Remap under Office add-in shortcut preferences</footer>
  </body>
</html>
`;
}

const check = process.argv.includes("--check");
const file = "public/shortcuts.html";
const path = new URL(`../${file}`, import.meta.url);
const expected = renderPage();

let current = "";
try {
  current = readFileSync(path, "utf8");
} catch {
  current = "";
}

if (current === expected) {
  process.stdout.write(`${file} up to date\n`);
} else if (check) {
  process.stderr.write(
    `${file} differs from public/shortcuts.json - run npm run shortcuts:build\n`,
  );
  process.exit(1);
} else {
  const tmp = new URL("../.shortcuts.html.tmp", import.meta.url);
  writeFileSync(tmp, expected);
  renameSync(tmp, path);
  process.stdout.write(`wrote ${file}\n`);
}
