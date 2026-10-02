// The generator's committed output, under test: every shortcut listed exactly
// once, the version present, and the npm wiring gates.test.ts's own style
// checks for the other generators (npm run check, npm run build). Reads the
// shipped public/shortcuts.html the same way copy.test.ts reads the shipped
// panes, rather than re-running the generator, so a stale commit fails here.

import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import pkg from "../package.json" with { type: "json" };
import shortcutsSource from "../public/shortcuts.json" with { type: "json" };

const html = readFileSync(
  path.join(process.cwd(), "public/shortcuts.html"),
  "utf8",
);

describe("public/shortcuts.html", () => {
  it("is self-contained: no script tags, no external requests", () => {
    expect(html).not.toContain("<script");
    expect(html).not.toMatch(/https?:\/\//);
  });

  it("carries the brand header and the footer line", () => {
    expect(html).toContain("<h1>pls,fix keyboard shortcuts</h1>");
    expect(html).toContain("Remap under Office add-in shortcut preferences");
  });

  it("contains the version from package.json", () => {
    expect(html).toContain(pkg.version);
  });

  it("lists every shortcut in the source exactly once", () => {
    for (const action of shortcutsSource.actions) {
      const count = html.split(`<span>${action.name}</span>`).length - 1;
      expect(count, `label "${action.name}"`).toBe(1);
    }
    for (const binding of shortcutsSource.shortcuts) {
      const combo = binding.key.default;
      const count = html.split(`<kbd>${combo}</kbd>`).length - 1;
      expect(count, `combo "${combo}"`).toBe(1);
    }
  });

  it("orders the shortcuts the way shortcuts.json - and the README - do", () => {
    const positions = shortcutsSource.actions.map((action) =>
      html.indexOf(`<span>${action.name}</span>`),
    );
    for (const position of positions) expect(position).toBeGreaterThan(-1);
    const sorted = [...positions].sort((a, b) => a - b);
    expect(positions).toEqual(sorted);
  });

  it("names no shortcut the source no longer ships", () => {
    const rows = html.match(/<span>[^<]*<\/span>/g) ?? [];
    const known = new Set(
      shortcutsSource.actions.map((action) => `<span>${action.name}</span>`),
    );
    for (const row of rows) expect(known.has(row), row).toBe(true);
  });
});

describe("npm run check and npm run build", () => {
  const scripts: Record<string, string> = pkg.scripts;

  it("defines shortcuts:build and shortcuts:check against the generator", () => {
    expect(scripts["shortcuts:build"]).toBe(
      "tsx scripts/build-shortcuts-page.ts",
    );
    expect(scripts["shortcuts:check"]).toBe(
      "tsx scripts/build-shortcuts-page.ts --check",
    );
  });

  it("runs shortcuts:check right after functions:check in npm run check", () => {
    const check = scripts.check!;
    expect(check).toContain(
      "npm run functions:check && npm run shortcuts:check",
    );
  });

  it("regenerates the page before vite build, the way functions:js runs", () => {
    expect(scripts.build).toContain(
      "tsx scripts/build-shortcuts-page.ts && npm run functions:js && vite build",
    );
  });
});
