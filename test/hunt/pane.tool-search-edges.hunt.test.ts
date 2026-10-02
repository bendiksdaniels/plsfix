// @vitest-environment jsdom
// Uncovered edges of the "Find a tool" search: runEntry's re-check of a
// button whose state moved between render and the click landing (a slow
// press, or busy elsewhere), handleListClick's guards against a click that
// did not land on a row at all, handleSlash ignoring every key but "/", and
// the item-1 sweep's odd values (unicode, empty, one letter) against the
// live wiring rather than the pure ranker (already covered in
// src/tool-search.test.ts).

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";

function paneRoot(): Document {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  return document;
}

async function load() {
  vi.resetModules();
  paneRoot();
  return import("../../src/pane/tool-search");
}

function input(): HTMLInputElement {
  return document.getElementById("tool-search") as HTMLInputElement;
}

function results(): HTMLElement {
  return document.getElementById("tool-search-results") as HTMLElement;
}

function type(value: string): void {
  const el = input();
  el.value = value;
  el.dispatchEvent(new Event("input"));
}

function press(key: string, target: Element = input()): void {
  target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true }));
}

describe("runEntry re-checks the button, never trusts render time", () => {
  it("a button disabled after render, before Enter lands, does nothing", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    type("export");
    expect(results().hidden).toBe(false);
    const button = document.getElementById(
      "export-selection",
    ) as HTMLButtonElement;
    const onClick = vi.fn();
    button.addEventListener("click", onClick);
    button.disabled = true; // the slow press: state moved on since render

    expect(() => press("Enter")).not.toThrow();
    expect(onClick).not.toHaveBeenCalled();
    // Nothing else about the box moved either: a no-op, not a half-run.
    expect(input().value).toBe("export");
    expect(results().hidden).toBe(false);
  });

  it("a row whose panel hid for its own reason (not just an inactive tab) does nothing on click", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    // New project's Create button is excluded from results while its own
    // prompt is hidden (existing coverage); this proves the SAME re-check
    // on the click path once a row IS showing and then its ancestor hides.
    type("undo");
    const row = results().querySelector("li")!;
    const button = document.querySelector(
      `[data-action="${row.dataset.toolAction}"]`,
    ) as HTMLButtonElement;
    const onClick = vi.fn();
    button.addEventListener("click", onClick);
    button.parentElement!.hidden = true;

    row.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(onClick).not.toHaveBeenCalled();
  });
});

describe("handleListClick's guards", () => {
  it("a click landing on a bare text node inside the list does nothing", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);
    type("export");

    const text = document.createTextNode("stray text");
    results().append(text);
    expect(() =>
      text.dispatchEvent(new MouseEvent("click", { bubbles: true })),
    ).not.toThrow();
  });

  it("a click on the list's own background, not a row, does nothing", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);
    type("export");

    const button = document.getElementById(
      "export-selection",
    ) as HTMLButtonElement;
    const onClick = vi.fn();
    button.addEventListener("click", onClick);

    results().dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(onClick).not.toHaveBeenCalled();
  });
});

describe('handleSlash ignores every key but "/"', () => {
  it("a different key from elsewhere in the pane leaves focus alone", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    const other = document.getElementById(
      "refresh-selection",
    ) as HTMLButtonElement;
    other.focus();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "k" }));
    expect(document.activeElement).toBe(other);
  });
});

describe("odd search values against the live wiring", () => {
  it("a unicode query never crashes and shows no matches", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    expect(() => type("Ünïcode 検索 💥")).not.toThrow();
    expect(results().hidden).toBe(true);
  });

  it("a single letter renders without throwing, matches or not", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    expect(() => type("z")).not.toThrow();
    // Whatever it finds (or does not), the list stays a valid ARIA listbox.
    expect(results().hidden).toBe(results().children.length === 0);
  });

  it("clearing back to empty hides the list again", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    type("export");
    expect(results().hidden).toBe(false);
    type("");
    expect(results().hidden).toBe(true);
  });
});

describe("cataloguing a synthetic malformed root", () => {
  it("a catalogued button outside any tabpanel reads an empty tab, not a crash", async () => {
    const { toolEntries } = await load();
    const stray = document.createElement("button");
    stray.dataset.action = "stray-action";
    stray.textContent = "Stray";
    document.body.append(stray); // not inside any [role=tabpanel]

    const entries = toolEntries(document);
    const entry = entries.find((e) => e.action === "stray-action");
    expect(entry).toMatchObject({ tab: "", label: "Stray" });
  });

  it("a help section whose key is unknown to HELP contributes no sentences", async () => {
    const { toolEntries } = await load();
    const section = document.createElement("section");
    section.dataset.help = "nonexistent-section-key";
    const heading = document.createElement("h3");
    heading.className = "section-heading";
    const button = document.createElement("button");
    button.id = "nonexistent-button";
    button.textContent = "Nonexistent";
    section.append(heading, button);
    document.body.append(section);

    expect(() => toolEntries(document)).not.toThrow();
    const entry = toolEntries(document).find(
      (e) => e.action === "nonexistent-button",
    );
    // No data-action and no known help sentence: never catalogued at all.
    expect(entry).toBeUndefined();
  });
});
