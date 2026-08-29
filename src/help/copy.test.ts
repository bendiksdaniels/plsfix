// @vitest-environment jsdom
// The coverage gate for the help copy. Parses the shipped taskpane.html and
// pptpane.html, so a button added, renamed or deleted fails here rather than
// leaving a silent hole in the "?" card. Both directions are checked: no
// button without a sentence, and no sentence naming a button that is gone.

import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { HELP, HELP_MAX_CHARS, helpSections, type HelpSection } from "./copy";

const PANES = ["taskpane.html", "pptpane.html"] as const;

// The jsdom environment turns import.meta.url into an http URL, so the panes
// are read from the vitest root instead - which is this repo.
function sectionsOf(pane: string): HelpSection[] {
  const html = readFileSync(path.join(process.cwd(), pane), "utf8");
  return helpSections(new DOMParser().parseFromString(html, "text/html"));
}

const PANE_SECTIONS = new Map(PANES.map((pane) => [pane, sectionsOf(pane)]));

function everySection(): { pane: string; section: HelpSection }[] {
  const all: { pane: string; section: HelpSection }[] = [];
  for (const [pane, sections] of PANE_SECTIONS) {
    for (const section of sections) all.push({ pane, section });
  }
  return all;
}

describe("help copy", () => {
  it("finds the sections of both panes", () => {
    expect(PANE_SECTIONS.get("taskpane.html")!.length).toBeGreaterThan(10);
    expect(PANE_SECTIONS.get("pptpane.html")!.length).toBeGreaterThan(2);
  });

  // One key means one section: the two panes both call a section "Linked
  // objects", and the deck's one carries data-help to stay apart from Excel's.
  it("keys every section apart across the two panes", () => {
    const seen = new Map<string, string>();
    const clashes: string[] = [];
    for (const { pane, section } of everySection()) {
      const first = seen.get(section.key);
      if (first !== undefined)
        clashes.push(`${section.key}: ${first}, ${pane}`);
      else seen.set(section.key, pane);
    }
    expect(clashes).toEqual([]);
  });

  it("has a line on what every section is for", () => {
    const missing = everySection()
      .filter(({ section }) => HELP[section.key] === undefined)
      .map(({ pane, section }) => `${pane} > ${section.key}`);
    expect(missing).toEqual([]);
  });

  it("has a sentence for every button in every section", () => {
    const missing: string[] = [];
    for (const { pane, section } of everySection()) {
      const copy = HELP[section.key];
      if (copy === undefined) continue;
      for (const button of section.buttons) {
        if (copy.buttons[button.key] === undefined) {
          missing.push(`${pane} > ${section.key} > ${button.key}`);
        }
      }
    }
    expect(missing).toEqual([]);
  });

  it("names no button that the panes no longer ship", () => {
    const stale: string[] = [];
    const covered = new Set<string>();
    for (const { section } of everySection()) {
      covered.add(section.key);
      const copy = HELP[section.key];
      if (copy === undefined) continue;
      const keys = new Set(section.buttons.map((button) => button.key));
      for (const key of Object.keys(copy.buttons)) {
        if (!keys.has(key)) stale.push(`${section.key} > ${key}`);
      }
    }
    for (const key of Object.keys(HELP)) {
      if (!covered.has(key)) stale.push(`${key} (no such section)`);
    }
    expect(stale).toEqual([]);
  });

  it("keeps every sentence short enough to read in a card", () => {
    const long: string[] = [];
    for (const [key, copy] of Object.entries(HELP)) {
      if (copy.about.length > HELP_MAX_CHARS) {
        long.push(`${key} (about, ${String(copy.about.length)})`);
      }
      for (const [button, text] of Object.entries(copy.buttons)) {
        if (text.length > HELP_MAX_CHARS) {
          long.push(`${key} > ${button} (${String(text.length)})`);
        }
      }
    }
    expect(long).toEqual([]);
  });

  it("writes plain sentences: no em dash, every one closed", () => {
    const bad: string[] = [];
    for (const [key, copy] of Object.entries(HELP)) {
      for (const [button, text] of [
        ["about", copy.about] as const,
        ...Object.entries(copy.buttons),
      ]) {
        if (text.includes("—")) bad.push(`${key} > ${button} (em dash)`);
        if (!text.endsWith(".")) bad.push(`${key} > ${button} (no full stop)`);
      }
    }
    expect(bad).toEqual([]);
  });
});
