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
});
