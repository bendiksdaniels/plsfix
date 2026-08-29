// Sample data for the UX gate: the rows, inbox items and option labels the
// panes are seeded with before they are measured. Deliberately awkward -
// long labels, a missing source, a stale push, the Latvian and Russian
// language strings - because a layout only breaks on the worst real content.

export const EXCEL_LINK_ROWS = [
  {
    entry: {
      id: "a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1",
      kind: "range",
      anchor: "PLSFIX_LINK_a1a1a1a1",
      label: "Revenue bridge FY25-FY26, EUR '000",
      token: "tok-1",
      createdAt: "2026-08-01T09:00:00Z",
      lastPushedAt: "2026-08-27T14:32:00Z",
      rev: 3,
    },
    source: "ok",
  },
  {
    entry: {
      id: "b2b2b2b2b2b2b2b2b2b2b2b2b2b2b2b2",
      kind: "chart",
      anchor: "PLSFIX_LINK_b2b2b2b2",
      label:
        "EBITDA margin trend, quarterly, all segments combined and restated",
      token: "tok-2",
      createdAt: "2026-07-15T09:00:00Z",
      lastPushedAt: null,
      rev: 1,
    },
    source: "ok",
  },
  {
    entry: {
      id: "c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3",
      kind: "range",
      anchor: "PLSFIX_LINK_c3c3c3c3",
      label: "Net debt bridge",
      token: "tok-3",
      createdAt: "2026-06-01T09:00:00Z",
      lastPushedAt: "2026-06-02T09:00:00Z",
      rev: 1,
    },
    source: "missing",
  },
  {
    entry: {
      id: "d4d4d4d4d4d4d4d4d4d4d4d4d4d4d4d4",
      kind: "chart",
      anchor: "PLSFIX_LINK_d4d4d4d4",
      label: "Segment revenue mix",
      token: "tok-4",
      createdAt: "2026-08-20T09:00:00Z",
      lastPushedAt: "2026-08-28T08:00:00Z",
      rev: 2,
    },
    source: "ok",
  },
];

export const CHART_OPTIONS = [
  "Revenue bridge FY25-FY26 by segment and region",
  "EBITDA margin trend",
  "Net debt waterfall (restated, IFRS 16 included)",
];

export function pptLinkRows(nowSeconds) {
  return [
    {
      key: "k1",
      slide: 3,
      label: "Revenue bridge FY25-FY26, EUR '000, full detail",
      source: "Q4 2026 Model - Consolidated Group View.xlsx",
      status: "current",
      pushedAt: nowSeconds - 3600,
      selected: true,
    },
    {
      key: "k2",
      slide: 12,
      label: "EBITDA margin trend",
      source: "Model.xlsx",
      status: "updateAvailable",
      pushedAt: nowSeconds - 90000,
      selected: false,
    },
    {
      key: "k3",
      slide: 5,
      label: "Segment revenue mix chart",
      source: "VeryLongWorkbookNameThatMightOverflowTheColumn.xlsx",
      status: "missing",
      pushedAt: null,
      selected: false,
    },
    {
      key: "k4",
      slide: 21,
      label: "Net debt bridge",
      source: "Model.xlsx",
      status: "wrongKey",
      pushedAt: nowSeconds - 200000,
      selected: false,
    },
  ];
}

export function inboxItems(nowMs) {
  return [
    {
      id: "i1",
      token: "tk1",
      kind: "range",
      label: "Revenue bridge FY25-FY26, EUR '000",
      src: {
        workbook: "Q4 2026 Model - Consolidated Group View.xlsx",
        sheet: "Summary",
        ref: "B2:H14",
        anchor: "PLSFIX_LINK_i1",
      },
      createdAt: new Date(nowMs - 3_600_000).toISOString(),
    },
    {
      id: "i2",
      token: "tk2",
      kind: "chart",
      label: "EBITDA trend",
      src: {
        workbook: "Model.xlsx",
        sheet: "Charts",
        ref: "Chart 1",
        anchor: "PLSFIX_LINK_i2",
      },
      createdAt: new Date(nowMs - 172_800_000).toISOString(),
    },
  ];
}

export const CHOOSER_OPTIONS = [
  "Revenue bridge FY25-FY26, EUR '000 - Q4 2026 Model - Consolidated Group View.xlsx",
  "EBITDA trend - Model.xlsx",
];

export const TOAST_TEXT =
  "Could not push 3 of 4 links: the relay refused the workspace key";
