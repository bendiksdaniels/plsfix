// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { installFirstRun } from "./first-run";

const KEY = "plsfix.firstRun.test.v1";

const CARD = `
  <section id="first-run">
    <p>1. Do a thing.</p>
    <div class="io-row">
      <button id="first-run-shortcuts" type="button">Shortcut card</button>
      <button id="first-run-dismiss" type="button">Got it</button>
    </div>
  </section>`;

function section(): HTMLElement {
  return document.getElementById("first-run")!;
}

// A minimal working store, stubbed in for the environment's own localStorage:
// this Node/jsdom combination ships a global `localStorage` whose methods are
// all undefined outside a real browser, so the module under test - which only
// ever calls getItem/setItem - is exercised against a real, controllable one.
let store: Map<string, string>;

function workingStorage(): Pick<Storage, "getItem" | "setItem"> {
  return {
    getItem: (name) => store.get(name) ?? null,
    setItem: (name, value) => {
      store.set(name, value);
    },
  };
}

function throwingStorage(): Pick<Storage, "getItem" | "setItem"> {
  return {
    getItem: () => {
      throw new Error("storage blocked");
    },
    setItem: () => {
      throw new Error("storage blocked");
    },
  };
}

describe("installFirstRun", () => {
  beforeEach(() => {
    document.body.innerHTML = CARD;
    store = new Map();
    vi.stubGlobal("localStorage", workingStorage());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("shows the card when storage has not seen this key", () => {
    installFirstRun(document, KEY, "first-run");
    expect(section().hidden).toBe(false);
  });

  it("hides after Got it and stays hidden on the next install", () => {
    installFirstRun(document, KEY, "first-run");
    document.getElementById("first-run-dismiss")!.click();
    expect(section().hidden).toBe(true);

    // A fresh install, as a new page load would run it, reads the same key
    // back from the same underlying store.
    document.body.innerHTML = CARD;
    installFirstRun(document, KEY, "first-run");
    expect(section().hidden).toBe(true);
  });

  it("calls onShortcuts when the shortcuts button is clicked", () => {
    const onShortcuts = vi.fn();
    installFirstRun(document, KEY, "first-run", onShortcuts);
    document.getElementById("first-run-shortcuts")!.click();
    expect(onShortcuts).toHaveBeenCalledTimes(1);
  });

  it("leaves the shortcuts button dead when no callback is given", () => {
    installFirstRun(document, KEY, "first-run");
    expect(() =>
      document.getElementById("first-run-shortcuts")!.click(),
    ).not.toThrow();
  });

  it("still shows the card and still lets dismiss run when storage throws", () => {
    vi.stubGlobal("localStorage", throwingStorage());

    installFirstRun(document, KEY, "first-run");
    expect(section().hidden).toBe(false);

    document.getElementById("first-run-dismiss")!.click();
    expect(section().hidden).toBe(true);
  });

  it("does nothing when the section is not on the page", () => {
    document.body.innerHTML = "";
    expect(() => installFirstRun(document, KEY, "first-run")).not.toThrow();
  });
});
