// @vitest-environment jsdom
// Hunt pass 1: the PowerPoint pane's Settings tab, split out of
// ppt.paste-pane.hunt.test.ts (its own relay mock and boot, no paste box or
// crypto machinery needed) - the workspace-key box's own Enter handler and
// switching link-transport from the Settings select.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import type * as RelayModule from "../../src/link/relay";
import { TRANSPORT_STORAGE_KEY } from "../../src/link/transport-setting";
import {
  createWorkspace,
  WORKSPACE_STORAGE_KEY,
  type Workspace,
} from "../../src/link/workspace";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
} from "../fakeppt";
import { FakeRelay } from "../fakerelay";
import { settlePpt, trackPptBoot } from "../ppt-ready";

enableStrictLoadSemantics();

let relay: FakeRelay;
vi.mock("../../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return relay;
    },
  };
});

interface BootOptions {
  local?: boolean;
}

async function bootPpt(options: BootOptions = {}): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
  relay = new FakeRelay();
  const storage = new Map<string, string>();
  if (options.local) {
    storage.set(TRANSPORT_STORAGE_KEY, "local");
  } else {
    const workspace: Workspace = await createWorkspace({
      get: async () => null,
      set: async () => undefined,
      remove: async () => undefined,
    });
    storage.set(WORKSPACE_STORAGE_KEY, workspace.exportKey);
    storage.set(TRANSPORT_STORAGE_KEY, "relay");
  }
  installFakePpt({ slides: 3, storage });
  const booted = trackPptBoot();
  await import("../../src/ppt/main");
  await booted;
  await settlePpt();
}

afterEach(() => {
  uninstallFakePpt();
});

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

describe("the workspace-key box's own Enter handler", () => {
  it("Enter runs the same save as clicking the button", async () => {
    await bootPpt({ local: false });
    const box = document.getElementById("workspace-key") as HTMLInputElement;
    box.value = "";

    box.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    await settlePpt();

    expect(toastText()).toBe("Paste the link key from Excel first.");
  });

  it("any other key does nothing at all", async () => {
    await bootPpt({ local: false });
    const box = document.getElementById("workspace-key") as HTMLInputElement;
    box.value = "not a key";

    box.dispatchEvent(
      new KeyboardEvent("keydown", { key: "a", bubbles: true }),
    );
    await settlePpt();

    expect(toastText()).toBe("");
    expect(box.value).toBe("not a key");
  });
});

describe("switching link-transport from the Settings select", () => {
  function transportSelect(): HTMLSelectElement {
    return document.getElementById("link-transport") as HTMLSelectElement;
  }

  it("relay to local says links now travel by copy and paste", async () => {
    await bootPpt({ local: false });
    const select = transportSelect();
    expect(select.value).toBe("relay");

    select.value = "local";
    select.dispatchEvent(new Event("change", { bubbles: true }));
    await settlePpt();

    expect(toastText()).toBe(
      "Links now travel by copy and paste on this computer.",
    );
    expect(document.getElementById("paste-links-area")!.hidden).toBe(false);
  });

  it("local to relay says links now travel through the relay", async () => {
    await bootPpt({ local: true });
    const select = transportSelect();
    expect(select.value).toBe("local");

    select.value = "relay";
    select.dispatchEvent(new Event("change", { bubbles: true }));
    await settlePpt();

    expect(toastText()).toBe("Links now travel through the relay.");
    expect(document.getElementById("paste-links-area")!.hidden).toBe(true);
  });
});
