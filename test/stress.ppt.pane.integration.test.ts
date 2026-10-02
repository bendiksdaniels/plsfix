// @vitest-environment jsdom
// Stress pass on the PowerPoint pane's buttons: pressed twice, pressed at the
// wrong moment, and pressed while a batch is still in the host. Drives the
// shipped pptpane.html through src/ppt/main.ts (harness: stress.ppt.support).
// Invariant: one action runs at a time, every press ends in a sentence, and
// the pane always comes back usable.

import { afterEach, describe, expect, it, vi } from "vitest";
import {
  bootPane,
  button,
  click,
  deck,
  heldFetch,
  hostHelpers,
  linkRows,
  paneRelay,
  plant,
  pressEnterOnKeyField,
  settle,
  tickRow,
  toastText,
  waiting,
} from "./stress.ppt.support";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import type * as RelayModule from "../src/link/relay";

enableStrictLoadSemantics();

vi.mock("../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return paneRelay();
    },
  };
});

afterEach(() => {
  uninstallFakePpt();
  vi.restoreAllMocks();
});

describe("a second action while the first is still running", () => {
  it("refuses the link key's Enter mid-batch and keeps the pane busy", async () => {
    await bootPane();
    await plant();
    click("refresh-links");
    await settle();
    const release = heldFetch();

    click("update-all");
    await settle(3);
    expect(button("update-all").disabled).toBe(true);

    // The field is not a button, so setBusy never disabled it: pressing Enter
    // here used to start a second flow whose own finish re-enabled every
    // button while the batch was still in the host.
    pressEnterOnKeyField();
    await settle();

    expect(toastText()).toBe("Wait for the last action to finish.");
    expect(button("update-all").disabled).toBe(true);
    expect(button("break-selected").disabled).toBe(true);

    release();
    await settle();
    expect(toastText()).toBe("1 up to date");
    expect(button("update-all").disabled).toBe(false);
  });

  it("refuses a paste mid-batch and processes nothing from it", async () => {
    await bootPane();
    await plant();
    click("refresh-links");
    await settle();
    const release = heldFetch();

    click("update-all");
    await settle(3);

    // The busy guard runs before pasteFromExcel even looks at the clipboard
    // data, so a bundle that would otherwise fail to decode never gets that
    // far: the refusal is "still busy", never a decode error.
    const box = document.getElementById("paste-links") as HTMLTextAreaElement;
    const event = new Event("paste", { bubbles: true, cancelable: true });
    Object.defineProperty(event, "clipboardData", {
      value: { getData: () => "" },
    });
    box.dispatchEvent(event);
    await settle();

    expect(toastText()).toBe("Wait for the last action to finish.");

    release();
    await settle();
    expect(toastText()).toBe("1 up to date");
  });

  it("disables every button while a batch runs and gives them all back", async () => {
    await bootPane();
    await plant();
    click("refresh-links");
    await settle();
    const release = heldFetch();

    click("update-all");
    await settle(3);
    const ids = [
      "refresh-links",
      "update-selected",
      "update-slide",
      "revert-selected",
      "break-selected",
      "go-to-slide",
      "refresh-inbox",
      "paste-latest-linked",
      "save-key",
      "forget-key",
      "clear-pasted-links",
      "align-objects",
    ];
    expect(ids.filter((id) => !button(id).disabled)).toEqual([]);

    release();
    await settle();
    expect(ids.filter((id) => button(id).disabled)).toEqual([]);
  });

  it("refuses a ribbon command while a batch is still in the host", async () => {
    await bootPane();
    await plant();
    // Two shapes on the selected slide, so Align left would have work to do.
    const slide = deck().slides[0]!;
    for (const left of [10, 90]) {
      deck().addShape(slide, { left, top: 10, width: 40, height: 40 });
    }
    hostHelpers().selectShapes(slide.shapes.map((shape) => shape.id));
    click("refresh-links");
    await settle();
    const release = heldFetch();

    click("update-all");
    await settle(3);
    // The ribbon shares this runtime: its press must meet the same latch as a
    // button's, not queue behind the batch and spend its own sync deadline.
    await hostHelpers().runCommand("PLSFIX_PPT_ALIGN_LEFT");
    await settle();

    expect(toastText()).toBe("Wait for the last action to finish.");
    expect(slide.shapes.map((shape) => shape.left)).toEqual([10, 90]);

    release();
    await settle();
    expect(toastText()).toBe("1 up to date");
    // And once the batch is done the same press works.
    await hostHelpers().runCommand("PLSFIX_PPT_ALIGN_LEFT");
    await settle();
    expect(toastText()).toBe("Aligned 2 objects left.");
  });

  it("inserts once when Paste latest linked is pressed twice", async () => {
    await bootPane();
    await waiting();
    click("refresh-inbox");
    await settle();

    click("paste-latest-linked");
    click("paste-latest-linked");
    await settle();

    expect(toastText()).toBe("Inserted Model!B4:F12.");
    expect(deck().slides[0]!.shapes).toHaveLength(1);
    expect(linkRows()).toHaveLength(1);
  });

  it("breaks once when Break is pressed twice", async () => {
    await bootPane();
    await plant();
    click("refresh-links");
    await settle();
    tickRow(0);

    click("break-selected");
    click("break-selected");
    await settle();

    expect(toastText()).toBe("1 link broken. The object stays on the slide.");
    expect(linkRows()).toHaveLength(0);
    expect(deck().slides[1]!.shapes).toHaveLength(1);
  });
});

describe("a button pressed at the wrong moment", () => {
  it("asks for a slide when nothing is selected in the deck", async () => {
    await bootPane();
    await waiting();
    click("refresh-inbox");
    await settle();
    hostHelpers().clearSelection();

    click("paste-latest-linked");
    await settle();

    expect(toastText()).toBe("insert Model!B4:F12: select a slide first.");
    expect(button("paste-latest-linked").disabled).toBe(false);
  });

  it("asks for a tick before Update selected, Revert, Break and Go to slide", async () => {
    await bootPane();
    await plant();
    click("refresh-links");
    await settle();

    for (const [id, sentence] of [
      ["update-selected", "Tick a link in the list first."],
      ["revert-selected", "Tick the rows to revert."],
      ["break-selected", "Tick a link in the list first."],
      ["go-to-slide", "Tick a link in the list first."],
    ] as [string, string][]) {
      // break-selected only arms on the first press (src/ui/confirm.ts); the
      // second is what actually runs and meets the empty selection.
      if (id === "break-selected") click(id);
      click(id);
      await settle();
      expect(toastText()).toBe(sentence);
      expect(button(id).disabled).toBe(false);
    }
  });

  it("says which slide is missing links when the wrong one is selected", async () => {
    await bootPane();
    await plant();
    click("refresh-links");
    await settle();
    hostHelpers().selectSlide(deck().slides[2]!.id);

    click("update-slide");
    await settle();

    expect(toastText()).toBe("No links on this slide");
    expect(button("update-slide").disabled).toBe(false);
  });

  // What this pins is the pane, not the host: the fake's setSelectedSlides
  // takes any id and office.js documents no error for a slide that is gone
  // (@types/office-js, Presentation.setSelectedSlides), so whether a real
  // PowerPoint throws is a launch-check row, not something to guess here.
  it("survives Go to slide on an id the deck no longer holds", async () => {
    await bootPane();
    await plant();
    click("refresh-links");
    await settle();
    tickRow(0);
    deck().slides.splice(1, 1);

    click("go-to-slide");
    await settle();

    expect(toastText()).not.toMatch(/ItemNotFound|Exception|undefined/);
    expect(button("go-to-slide").disabled).toBe(false);
  });

  it("keeps working after every refusal in a row", async () => {
    await bootPane();
    await plant();

    for (const id of [
      "update-selected",
      "revert-selected",
      "break-selected",
      "go-to-slide",
      "change-source",
    ]) {
      click(id);
      await settle();
      expect(toastText()).not.toBe("");
    }

    click("refresh-links");
    await settle();
    expect(toastText()).toBe("1 linked object.");
  });
});

describe("the Inbox tab coming into view", () => {
  it("reads the inbox when the tab is clicked, not only at boot", async () => {
    await bootPane();
    // Boot itself already read the inbox once, before this item existed.
    await waiting();

    click("tab-inbox");
    await settle();

    expect(
      document.querySelectorAll("#inbox-list button").length,
    ).toBeGreaterThan(0);
    expect(toastText()).toBe("1 waiting to insert.");
  });

  it("does nothing, no toast, when the pane is not paired", async () => {
    await bootPane(false);

    click("tab-inbox");
    await settle();

    expect(toastText()).toBe("");
  });
});

describe("a ribbon command showing the pane on error", () => {
  function showAsTaskpaneSpy(): ReturnType<typeof vi.spyOn> {
    return vi.spyOn(
      (globalThis as { Office: { addin: { showAsTaskpane: () => void } } })
        .Office.addin,
      "showAsTaskpane",
    );
  }

  it("shows the pane before toasting, so a closed pane's toast is not lost", async () => {
    await bootPane();
    const spy = showAsTaskpaneSpy();

    // Nothing selected: Swap needs exactly two objects.
    await hostHelpers().runCommand("PLSFIX_PPT_SWAP");
    await settle();

    expect(spy).toHaveBeenCalledOnce();
    expect(toastText()).toBe("Select exactly two objects to swap.");
  });

  it("does not show the pane when the command succeeds", async () => {
    await bootPane();
    const slide = deck().slides[0]!;
    const a = deck().addShape(slide, {
      left: 10,
      top: 10,
      width: 40,
      height: 40,
    });
    const b = deck().addShape(slide, {
      left: 90,
      top: 10,
      width: 40,
      height: 40,
    });
    hostHelpers().selectShapes([a.id, b.id]);
    const spy = showAsTaskpaneSpy();

    await hostHelpers().runCommand("PLSFIX_PPT_SWAP");
    await settle();

    expect(spy).not.toHaveBeenCalled();
    expect(toastText()).toBe("Swapped two objects.");
  });

  it("does not show the pane for a regular button's failure - it is already open", async () => {
    await bootPane();
    await plant();
    click("refresh-links");
    await settle();
    const spy = showAsTaskpaneSpy();

    // No row ticked.
    click("update-selected");
    await settle();

    expect(spy).not.toHaveBeenCalled();
    expect(toastText()).toBe("Tick a link in the list first.");
  });
});
