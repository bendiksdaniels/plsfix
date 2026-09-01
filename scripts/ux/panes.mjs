// What the UX gate walks: the two panes, and per pane the states a user can
// actually reach - a filled link list, an empty one, the chart picker, the
// "Change source" chooser, a toast. Each state seeds the real render
// functions through a Vite-served dynamic import, so the DOM under test is
// the DOM the pane builds, never a hand-written copy of it.

import {
  CHART_OPTIONS,
  CHOOSER_OPTIONS,
  EXCEL_LINK_ROWS,
  TOAST_TEXT,
  inboxItems,
  pptLinkRows,
} from "./fixtures.mjs";

async function setConnection(page, text, cls) {
  await page.evaluate(
    ([label, className]) => {
      const el = document.getElementById("connection-status");
      if (el) {
        el.textContent = label;
        el.className = `connection ${className}`;
      }
    },
    [text, cls],
  );
}

async function seedExcelLinks(page, rows) {
  await page.evaluate(async (data) => {
    const mod = await import("/src/pane/links-list.ts");
    const tbody = document.getElementById("workbook-links");
    if (!tbody) throw new Error("#workbook-links missing");
    const selected = new Set(data.length > 0 ? [data[0].entry.id] : []);
    mod.renderWorkbookLinks(tbody, data, selected, () => undefined);
  }, rows);
}

async function fillSelect(page, id, labels, show) {
  await page.evaluate(
    ([selectId, options, visible]) => {
      const el = document.getElementById(selectId);
      if (!el) throw new Error(`#${selectId} missing`);
      el.replaceChildren();
      for (const label of options) {
        const option = document.createElement("option");
        option.textContent = label;
        el.append(option);
      }
      el.hidden = !visible;
    },
    [id, labels, show],
  );
}

async function showToast(page) {
  await page.evaluate(async (text) => {
    const mod = await import("/src/ui/toast.ts");
    const node = document.getElementById("toast");
    if (!node) throw new Error("#toast missing");
    mod.createToast(node, 600_000).show(text, "error", "stack trace here");
  }, TOAST_TEXT);
}

async function seedPptRows(page, rows) {
  await page.evaluate(async (data) => {
    const mod = await import("/src/ppt/views.ts");
    const body = document.getElementById("link-rows");
    if (!body) throw new Error("#link-rows missing");
    mod.renderLinkRows(body, data, () => undefined);
    const empty = document.getElementById("links-empty");
    if (empty) empty.hidden = data.length > 0;
  }, rows);
}

async function seedInbox(page, items, unpaired) {
  await page.evaluate(
    async ([data, showUnpaired]) => {
      const mod = await import("/src/ppt/views.ts");
      const list = document.getElementById("inbox-list");
      if (!list) throw new Error("#inbox-list missing");
      mod.renderInbox(list, data, () => undefined);
      list.hidden = false;
      const hint = document.getElementById("inbox-unpaired");
      if (hint) hint.hidden = !showUnpaired;
    },
    [items, unpaired],
  );
}

async function setDisabled(page, ids, off) {
  await page.evaluate(
    ([targets, disabled]) => {
      for (const id of targets) {
        const el = document.getElementById(id);
        if (el instanceof HTMLButtonElement) el.disabled = disabled;
      }
    },
    [ids, off],
  );
}

async function toggleChooser(page, open) {
  await page.evaluate((show) => {
    const el = document.getElementById("change-source-chooser");
    if (el) el.hidden = !show;
  }, open);
}

const TASKPANE_STATES = [
  {
    slug: "filled",
    focus: true,
    apply: async (page) => {
      await setConnection(page, "Excel connected", "ready");
      await seedExcelLinks(page, EXCEL_LINK_ROWS);
      await fillSelect(page, "export-chart-pick", CHART_OPTIONS, false);
    },
  },
  {
    slug: "empty",
    tabs: ["tools", "workbook", "links"],
    apply: async (page) => {
      await setConnection(page, "Connecting", "waiting");
      await seedExcelLinks(page, []);
      await fillSelect(page, "export-chart-pick", [], false);
    },
  },
  {
    slug: "chartlist",
    tabs: ["links"],
    apply: async (page) => {
      await setConnection(page, "Excel connected", "ready");
      await seedExcelLinks(page, EXCEL_LINK_ROWS);
      await fillSelect(page, "export-chart-pick", CHART_OPTIONS, true);
    },
  },
  {
    slug: "toast",
    tabs: ["tools"],
    apply: async (page) => {
      await setConnection(page, "Excel required", "error");
      await showToast(page);
    },
  },
];

const PPT_STATES = [
  {
    slug: "filled",
    focus: true,
    apply: async (page) => {
      await setConnection(page, "PowerPoint connected", "ready");
      await seedPptRows(page, pptLinkRows(Date.now() / 1000));
      await seedInbox(page, inboxItems(Date.now()), false);
      await toggleChooser(page, false);
      await setDisabled(page, ["change-source"], false);
    },
  },
  {
    slug: "empty",
    tabs: ["links", "inbox", "tools"],
    apply: async (page) => {
      await setConnection(page, "Connecting", "waiting");
      await seedPptRows(page, []);
      await seedInbox(page, [], true);
      await setDisabled(page, ["change-source", "revert-selected"], true);
    },
  },
  {
    slug: "chooser",
    tabs: ["links"],
    apply: async (page) => {
      await setConnection(page, "PowerPoint connected", "ready");
      await seedPptRows(page, pptLinkRows(Date.now() / 1000));
      await toggleChooser(page, true);
      await fillSelect(page, "change-source-list", CHOOSER_OPTIONS, true);
    },
  },
  {
    slug: "toast",
    tabs: ["links"],
    apply: async (page) => {
      await setConnection(page, "PowerPoint required", "error");
      await toggleChooser(page, false);
      await showToast(page);
    },
  },
];

export const PANES = [
  { file: "taskpane.html", slug: "taskpane", states: TASKPANE_STATES },
  { file: "pptpane.html", slug: "pptpane", states: PPT_STATES },
];
