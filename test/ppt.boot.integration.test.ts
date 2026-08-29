// @vitest-environment jsdom
// The PowerPoint pane's boot gate: every button is wired at first paint, so a
// click before Office.onReady has confirmed PowerPoint and PowerPointApi 1.5
// must say the host is not connected rather than reach PowerPoint.run. Driven
// through the real pptpane.html markup and the fake deck.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { installFakePpt, uninstallFakePpt } from "./fakeppt";

function pane(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
}

// The pane boots on import, so the markup and the host go in first.
async function boot(supported: boolean): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  pane();
  installFakePpt({ slides: 1, isSetSupported: () => supported });
  await import("../src/ppt/main");
  await new Promise((resolve) => setTimeout(resolve, 0));
}

function click(id: string): void {
  document.querySelector<HTMLButtonElement>(`#${id}`)?.click();
}

async function settle(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

function badge(): string {
  return document.getElementById("connection-status")?.textContent ?? "";
}

afterEach(() => {
  uninstallFakePpt();
});

describe("a host the pane rejected", () => {
  it("says so on the badge and on every button", async () => {
    await boot(false);
    expect(badge()).toBe("PowerPoint 2021 / Microsoft 365 required");

    click("refresh-links");
    await settle();
    expect(toastText()).toBe("PowerPoint is not connected.");
    expect(document.getElementById("toast")?.className).toContain("error");

    click("update-all");
    await settle();
    expect(toastText()).toBe("PowerPoint is not connected.");
  });

  it("refuses the Enter key on the link key field too", async () => {
    await boot(false);
    const field = document.getElementById("workspace-key") as HTMLInputElement;
    field.value = "not a key";
    field.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    await settle();

    expect(toastText()).toBe("PowerPoint is not connected.");
    expect(field.value).toBe("not a key");
  });
});

describe("a host the pane accepted", () => {
  it("runs the same button", async () => {
    await boot(true);
    expect(badge()).toBe("PowerPoint connected");

    click("refresh-links");
    await settle();
    expect(toastText()).toBe("No linked objects in this deck.");
  });

  // "Revert last update" acts on ticks, so with none it must say which ticks
  // it wants rather than reverting the whole deck by surprise.
  it("asks for ticks before reverting anything", async () => {
    await boot(true);

    click("revert-selected");
    await settle();

    expect(toastText()).toBe("Tick the rows to revert.");
  });
});
