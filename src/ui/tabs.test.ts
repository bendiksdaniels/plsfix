// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { installTabs } from "./tabs";

describe("installTabs", () => {
  it("shows exactly one panel and marks its tab selected", () => {
    document.body.innerHTML = `
      <nav id="bar"><button role="tab" id="t1" aria-controls="p1" aria-selected="true" class="tab active"></button>
      <button role="tab" id="t2" aria-controls="p2" aria-selected="false" class="tab"></button></nav>
      <div id="p1"></div><div id="p2" hidden></div>`;
    const tabs = installTabs(document.getElementById("bar")!);
    tabs.activate("t2");
    expect(document.getElementById("p1")!.hidden).toBe(true);
    expect(document.getElementById("p2")!.hidden).toBe(false);
    expect(document.getElementById("t2")!.getAttribute("aria-selected")).toBe(
      "true",
    );
    expect(document.getElementById("t1")!.classList.contains("active")).toBe(
      false,
    );
  });
  it("activates on a real click, not just via activate()", () => {
    document.body.innerHTML = `
      <nav id="bar"><button role="tab" id="t1" aria-controls="p1" aria-selected="true" class="tab active"></button>
      <button role="tab" id="t2" aria-controls="p2" aria-selected="false" class="tab"></button></nav>
      <div id="p1"></div><div id="p2" hidden></div>`;
    installTabs(document.getElementById("bar")!);
    document.getElementById("t2")!.click();
    expect(document.getElementById("p1")!.hidden).toBe(true);
    expect(document.getElementById("p2")!.hidden).toBe(false);
    expect(document.getElementById("t2")!.getAttribute("aria-selected")).toBe(
      "true",
    );
    expect(document.getElementById("t1")!.classList.contains("active")).toBe(
      false,
    );
  });

  function threeTabBar(): HTMLElement {
    document.body.innerHTML = `
      <nav id="bar">
        <button role="tab" id="t1" aria-controls="p1" aria-selected="true" class="tab active"></button>
        <button role="tab" id="t2" aria-controls="p2" aria-selected="false" class="tab"></button>
        <button role="tab" id="t3" aria-controls="p3" aria-selected="false" class="tab"></button>
      </nav>
      <div id="p1"></div><div id="p2" hidden></div><div id="p3" hidden></div>`;
    return document.getElementById("bar")!;
  }

  function press(tabId: string, key: string): void {
    document
      .getElementById(tabId)!
      .dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true }));
  }

  it("moves focus and activates the next tab on ArrowRight", () => {
    installTabs(threeTabBar());
    document.getElementById("t1")!.focus();
    press("t1", "ArrowRight");
    expect(document.activeElement?.id).toBe("t2");
    expect(document.getElementById("t2")!.getAttribute("aria-selected")).toBe(
      "true",
    );
    expect(document.getElementById("p2")!.hidden).toBe(false);
    expect(document.getElementById("p1")!.hidden).toBe(true);
  });

  it("wraps from the last tab to the first on ArrowRight", () => {
    installTabs(threeTabBar());
    document.getElementById("t3")!.focus();
    press("t3", "ArrowRight");
    expect(document.activeElement?.id).toBe("t1");
  });

  it("wraps from the first tab to the last on ArrowLeft", () => {
    installTabs(threeTabBar());
    document.getElementById("t1")!.focus();
    press("t1", "ArrowLeft");
    expect(document.activeElement?.id).toBe("t3");
    expect(document.getElementById("t3")!.getAttribute("aria-selected")).toBe(
      "true",
    );
  });

  it("Home and End jump to the first and last tab", () => {
    installTabs(threeTabBar());
    document.getElementById("t2")!.focus();
    press("t2", "End");
    expect(document.activeElement?.id).toBe("t3");
    press("t3", "Home");
    expect(document.activeElement?.id).toBe("t1");
  });

  it("ArrowUp/ArrowDown work the same as Left/Right for a stacked bar", () => {
    installTabs(threeTabBar());
    document.getElementById("t1")!.focus();
    press("t1", "ArrowDown");
    expect(document.activeElement?.id).toBe("t2");
    press("t2", "ArrowUp");
    expect(document.activeElement?.id).toBe("t1");
  });

  it("leaves every other key alone, Tab included", () => {
    installTabs(threeTabBar());
    document.getElementById("t1")!.focus();
    press("t1", "Tab");
    // No preventDefault, no activation change - the browser's own Tab order
    // is left to do its job.
    expect(document.getElementById("t1")!.getAttribute("aria-selected")).toBe(
      "true",
    );
  });

  it("ignores a keydown that did not originate on a tab", () => {
    const bar = threeTabBar();
    installTabs(bar);
    expect(() =>
      bar.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight" })),
    ).not.toThrow();
    expect(document.getElementById("t1")!.getAttribute("aria-selected")).toBe(
      "true",
    );
  });

  // The ARIA tabs pattern: the strip is one Tab stop, and the arrows move
  // inside it. Without this a four-tab pane costs four presses to walk past.
  function tabStops(): number[] {
    return ["t1", "t2", "t3"].map(
      (id) => document.getElementById(id)!.tabIndex,
    );
  }

  it("makes the active tab the strip's only Tab stop on install", () => {
    installTabs(threeTabBar());
    expect(tabStops()).toEqual([0, -1, -1]);
  });

  it("moves the Tab stop with the activation, however it happened", () => {
    const tabs = installTabs(threeTabBar());
    tabs.activate("t3");
    expect(tabStops()).toEqual([-1, -1, 0]);
    document.getElementById("t2")!.click();
    expect(tabStops()).toEqual([-1, 0, -1]);
    document.getElementById("t2")!.focus();
    press("t2", "ArrowRight");
    expect(tabStops()).toEqual([-1, -1, 0]);
  });

  it("makes the first tab the stop when the markup selects none", () => {
    document.body.innerHTML = `
      <nav id="bar">
        <button role="tab" id="t1" aria-controls="p1" class="tab"></button>
        <button role="tab" id="t2" aria-controls="p2" class="tab"></button>
        <button role="tab" id="t3" aria-controls="p3" class="tab"></button>
      </nav>
      <div id="p1"></div><div id="p2"></div><div id="p3"></div>`;
    installTabs(document.getElementById("bar")!);
    expect(tabStops()).toEqual([0, -1, -1]);
  });

  it("survives a bar with no tabs in it at all", () => {
    document.body.innerHTML = `<nav id="bar"></nav>`;
    expect(() => installTabs(document.getElementById("bar")!)).not.toThrow();
  });

  // The Excel pane hangs the sheet explorer's refresh on the Workbook tab's
  // own click, so an arrow-key activation has to be a click, not a quiet
  // aria-selected flip, or the tab opens with last session's sheets.
  it("runs an arrow-key activation through the click every hook is on", () => {
    installTabs(threeTabBar());
    const hooked: string[] = [];
    document
      .getElementById("t2")!
      .addEventListener("click", () => hooked.push("t2"));
    document.getElementById("t1")!.focus();
    press("t1", "ArrowRight");
    expect(hooked).toEqual(["t2"]);
    expect(document.getElementById("p2")!.hidden).toBe(false);
    expect(document.activeElement?.id).toBe("t2");
  });

  it("fires that click once per activation, not once per tab", () => {
    installTabs(threeTabBar());
    const hooked: string[] = [];
    for (const id of ["t1", "t2", "t3"]) {
      document
        .getElementById(id)!
        .addEventListener("click", () => hooked.push(id));
    }
    document.getElementById("t1")!.focus();
    press("t1", "End");
    press("t3", "Home");
    expect(hooked).toEqual(["t3", "t1"]);
  });
});
