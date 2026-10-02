// The six ready calculation blocks as pure data, one grid of typed cells each:
// a debt schedule, a DCF, an NPV/IRR block, working-capital days, a two-way
// sensitivity grid and an EBITDA bridge. Owns every number, label and formula a
// template writes; the invariant is that no placeholder points outside its own
// block, which `templates.test.ts` re-checks for every grid.

import {
  axis,
  blank,
  blankRow,
  build,
  calc,
  head,
  input,
  label,
  numberCell,
  pad,
  result,
  type Template,
  type TemplateCell,
  title,
  total,
} from "./template-cells";

// ---------------------------------------------------------------------------
// 1. Debt schedule (annuity)
// ---------------------------------------------------------------------------

const DEBT_COLS = 6;
const DEBT_PERIODS = 10;
// The row the first period sits on: title, four inputs, two derived, a gap, a header.
const DEBT_FIRST = 9;

function debtPeriod(index: number): TemplateCell[] {
  const row = DEBT_FIRST + index;
  const above = row - 1;
  return [
    index === 0
      ? numberCell(1, "input", "whole")
      : calc(`={r+${above},c}+1`, "whole"),
    calc(index === 0 ? "={r+1,c+1}" : `={r+${above},c+5}`, "currency"),
    calc("={r+6,c+1}", "currency"),
    calc(`={r+${row},c+1}*{r+2,c+1}/{r+4,c+1}`, "currency"),
    calc(`={r+${row},c+2}-{r+${row},c+3}`, "currency"),
    calc(`={r+${row},c+1}-{r+${row},c+4}`, "currency"),
  ];
}

// Title, the four inputs, the two numbers they derive, a gap and the header.
function debtInputs(): TemplateCell[][] {
  return [
    pad([title("Debt schedule (annuity)")], DEBT_COLS, "title"),
    pad([label("Principal"), input(1_000_000, "currency")], DEBT_COLS),
    pad([label("Annual rate"), input(0.06, "percent")], DEBT_COLS),
    pad([label("Years"), input(5, "whole")], DEBT_COLS),
    pad([label("Payments per year"), input(2, "whole")], DEBT_COLS),
    pad([label("Periods"), calc("={r+3,c+1}*{r+4,c+1}", "whole")], DEBT_COLS),
    pad(
      [
        label("Payment"),
        calc("=PMT({r+2,c+1}/{r+4,c+1},{r+5,c+1},-{r+1,c+1})", "currency"),
      ],
      DEBT_COLS,
    ),
    blankRow(DEBT_COLS),
    ["Period", "Opening", "Payment", "Interest", "Principal", "Closing"].map(
      (name) => head(name),
    ),
  ];
}

// Only the three flow columns total; a balance column has no meaningful sum.
function debtTotals(): TemplateCell[] {
  const last = DEBT_FIRST + DEBT_PERIODS - 1;
  const sum = (column: number): TemplateCell =>
    total(
      `=SUM({r+${DEBT_FIRST},c+${column}}:{r+${last},c+${column}})`,
      "currency",
    );
  return [
    result("Total"),
    blank("result"),
    sum(2),
    sum(3),
    sum(4),
    blank("result"),
  ];
}

export function debtSchedule(): Template {
  return build(
    "debt-schedule",
    "Debt schedule (annuity)",
    "Principal, rate, years and payments per year into a level-payment schedule with interest, principal and closing balance per period.",
    [
      ...debtInputs(),
      ...Array.from({ length: DEBT_PERIODS }, (_unused, index) =>
        debtPeriod(index),
      ),
      debtTotals(),
    ],
  );
}

// ---------------------------------------------------------------------------
// 2. DCF block
// ---------------------------------------------------------------------------

const DCF_COLS = 6;
const DCF_FCF = [100_000, 110_000, 120_000, 130_000, 140_000];
const DCF_LAST = DCF_FCF.length;

function dcfInputs(): TemplateCell[][] {
  return [
    pad([title("DCF valuation")], DCF_COLS, "title"),
    pad([label("WACC"), input(0.09, "percent")], DCF_COLS),
    pad([label("Terminal growth"), input(0.02, "percent")], DCF_COLS),
    blankRow(DCF_COLS),
  ];
}

// The forecast strip: the year header, the cash flows, the factor each year is
// discounted by and the present value the two of them make.
function dcfForecast(): TemplateCell[][] {
  const across = (make: (index: number) => TemplateCell): TemplateCell[] =>
    DCF_FCF.map((_unused, index) => make(index));
  return [
    [
      head("Year"),
      ...across((index) => numberCell(index + 1, "header", "whole")),
    ],
    [label("Free cash flow"), ...DCF_FCF.map((v) => input(v, "currency"))],
    [
      label("Discount factor"),
      ...across((index) =>
        calc(`=1/(1+{r+1,c+1})^{r+4,c+${index + 1}}`, "percent"),
      ),
    ],
    [
      label("PV of FCF"),
      ...across((index) =>
        calc(`={r+5,c+${index + 1}}*{r+6,c+${index + 1}}`, "currency"),
      ),
    ],
    blankRow(DCF_COLS),
  ];
}

// Gordon growth off the last forecast year, discounted by the last factor.
function dcfValuation(): TemplateCell[][] {
  return [
    pad(
      [
        label("PV of forecast FCF"),
        calc(`=SUM({r+7,c+1}:{r+7,c+${DCF_LAST}})`, "currency"),
      ],
      DCF_COLS,
    ),
    pad(
      [
        label("Terminal value"),
        calc(
          `={r+5,c+${DCF_LAST}}*(1+{r+2,c+1})/({r+1,c+1}-{r+2,c+1})`,
          "currency",
        ),
      ],
      DCF_COLS,
    ),
    pad(
      [
        label("PV of terminal value"),
        calc(`={r+10,c+1}*{r+6,c+${DCF_LAST}}`, "currency"),
      ],
      DCF_COLS,
    ),
    pad(
      [result("Enterprise value"), total("={r+9,c+1}+{r+11,c+1}", "currency")],
      DCF_COLS,
      "result",
    ),
  ];
}

export function dcf(): Template {
  return build(
    "dcf",
    "DCF valuation",
    "WACC, terminal growth and five free cash flows into discount factors, present values, a Gordon terminal value and enterprise value.",
    [...dcfInputs(), ...dcfForecast(), ...dcfValuation()],
  );
}

// ---------------------------------------------------------------------------
// 3. NPV / IRR block
// ---------------------------------------------------------------------------

const NPV_COLS = 3;
const NPV_FLOWS = [
  80_000, 95_000, 105_000, 115_000, 125_000, 135_000, 145_000, 155_000,
];
// Period 0 carries the outlay; the eight flows follow it.
const NPV_FIRST = 4;
const NPV_LAST = NPV_FIRST + NPV_FLOWS.length;

function npvPeriod(index: number): TemplateCell[] {
  const row = NPV_FIRST + index + 1;
  const above = row - 1;
  return [
    calc(`={r+${above},c}+1`, "whole"),
    input(NPV_FLOWS[index] ?? 0, "currency"),
    calc(`={r+${above},c+2}+{r+${row},c+1}`, "currency"),
  ];
}

// Periods to recover the outlay, plus the fraction of the next flow it takes:
// the count of still-negative cumulative cells is the whole part, and INDEX
// reads the shortfall and the flow that closes it. INDEX runs off the end when
// the block never turns positive, which is the case IFERROR names.
function paybackFormula(): string {
  const cumulative = `{r+${NPV_FIRST},c+2}:{r+${NPV_LAST},c+2}`;
  const flows = `{r+${NPV_FIRST},c+1}:{r+${NPV_LAST},c+1}`;
  const behind = `COUNTIF(${cumulative},"<0")`;
  return `=IFERROR((${behind}-1)-INDEX(${cumulative},${behind})/INDEX(${flows},${behind}+1),"n/a")`;
}

function npvHeader(): TemplateCell[][] {
  return [
    pad([title("NPV / IRR")], NPV_COLS, "title"),
    pad([label("Discount rate"), input(0.1, "percent")], NPV_COLS),
    blankRow(NPV_COLS),
    [head("Period"), head("Cash flow"), head("Cumulative")],
    [
      numberCell(0, "input", "whole"),
      input(-500_000, "currency"),
      calc(`={r+${NPV_FIRST},c+1}`, "currency"),
    ],
  ];
}

// NPV discounts the eight later flows and adds the outlay back at face value,
// the way Excel's own NPV wants it; IRR takes the outlay in with the rest.
function npvResults(): TemplateCell[][] {
  const npv = `=NPV({r+1,c+1},{r+${NPV_FIRST + 1},c+1}:{r+${NPV_LAST},c+1})+{r+${NPV_FIRST},c+1}`;
  const irr = `=IRR({r+${NPV_FIRST},c+1}:{r+${NPV_LAST},c+1})`;
  return [
    blankRow(NPV_COLS),
    pad([result("NPV"), total(npv, "currency")], NPV_COLS, "result"),
    pad([result("IRR"), total(irr, "percent")], NPV_COLS, "result"),
    pad(
      [result("Payback period"), total(paybackFormula(), "decimal")],
      NPV_COLS,
      "result",
    ),
  ];
}

export function npvIrr(): Template {
  return build(
    "npv-irr",
    "NPV / IRR",
    "An outlay and eight periodic cash flows into a cumulative column, NPV at your discount rate, IRR and a payback period.",
    [
      ...npvHeader(),
      ...Array.from({ length: NPV_FLOWS.length }, (_unused, index) =>
        npvPeriod(index),
      ),
      ...npvResults(),
    ],
  );
}

// ---------------------------------------------------------------------------
// 4. Working-capital days
// ---------------------------------------------------------------------------

export function workingCapital(): Template {
  return build(
    "working-capital",
    "Working capital days",
    "Revenue, COGS and the three working-capital balances into DSO, DIO, DPO and the cash conversion cycle.",
    [
      [title("Working capital days"), blank("title")],
      [label("Revenue"), input(5_000_000, "currency")],
      [label("COGS"), input(3_000_000, "currency")],
      [label("Receivables"), input(820_000, "currency")],
      [label("Inventory"), input(640_000, "currency")],
      [label("Payables"), input(510_000, "currency")],
      blankRow(2),
      [head("Metric"), head("Days")],
      [label("DSO"), calc("={r+3,c+1}/{r+1,c+1}*365", "decimal")],
      [label("DIO"), calc("={r+4,c+1}/{r+2,c+1}*365", "decimal")],
      [label("DPO"), calc("={r+5,c+1}/{r+2,c+1}*365", "decimal")],
      [
        result("Cash conversion cycle"),
        total("={r+8,c+1}+{r+9,c+1}-{r+10,c+1}", "decimal"),
      ],
    ],
  );
}

// ---------------------------------------------------------------------------
// 5. Two-way sensitivity grid
// ---------------------------------------------------------------------------

const SENS_STEPS = [-2, -1, 0, 1, 2];
const SENS_HEADER = 5;
const SENS_FIRST = SENS_HEADER + 1;

function sensitivityRow(index: number): TemplateCell[] {
  const row = SENS_FIRST + index;
  return [
    axis(`={r+2,c+1}*${SENS_STEPS[index] ?? 0}`),
    ...SENS_STEPS.map((_unused, column) =>
      calc(
        `={r+1,c+1}*(1+{r+${row},c})*(1+{r+${SENS_HEADER},c+${column + 1}})`,
        "currency",
      ),
    ),
  ];
}

export function sensitivity(): Template {
  const width = SENS_STEPS.length + 1;
  return build(
    "sensitivity",
    "Two-way sensitivity",
    "A base value and two drivers with their step sizes into a 5x5 grid of formulas, the axes stepping in both directions.",
    [
      pad([title("Two-way sensitivity")], width, "title"),
      pad([label("Base value"), input(1_000_000, "currency")], width),
      pad([label("Row driver step"), input(0.05, "percent")], width),
      pad([label("Column driver step"), input(0.1, "percent")], width),
      blankRow(width),
      [
        blank("header"),
        ...SENS_STEPS.map((step) => axis(`={r+3,c+1}*${step}`)),
      ],
      ...SENS_STEPS.map((_unused, index) => sensitivityRow(index)),
    ],
  );
}

// ---------------------------------------------------------------------------
// 6. EBITDA bridge table
// ---------------------------------------------------------------------------

const BRIDGE_STEPS: [string, number][] = [
  ["Price", 900_000],
  ["Volume", 600_000],
  ["Mix", -250_000],
  ["Cost inflation", -700_000],
  ["Other", 150_000],
];

export function ebitdaBridge(): Template {
  const lastStep = 1 + BRIDGE_STEPS.length;
  return build(
    "ebitda-bridge",
    "EBITDA bridge",
    "Opening EBITDA, five named steps and closing EBITDA in the two-column shape the waterfall chart reads, plus a check row.",
    [
      [title("EBITDA bridge"), blank("title")],
      [label("Opening EBITDA"), input(12_000_000, "currency")],
      ...BRIDGE_STEPS.map(([name, value]) => [
        label(name),
        input(value, "currency"),
      ]),
      [label("Closing EBITDA"), input(12_700_000, "currency")],
      [
        result("Check"),
        total(
          `={r+1,c+1}+SUM({r+2,c+1}:{r+${lastStep},c+1})-{r+${lastStep + 1},c+1}`,
          "currency",
        ),
      ],
    ],
  );
}
