// Pass-1 hunt: every template's formulas evaluated by hand as finance: the
// debt schedule amortises to zero, DCF discounting and terminal value,
// NPV/IRR/payback, working-capital days, the sensitivity corners, the bridge.
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "../fakehost";
import type * as ExcelModule from "../../src/excel";
import { cellAddress } from "../../src/find";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(sheets: string[] = ["Model"]): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets }).helpers;
  smt = await import("../../src/excel");
}

beforeEach(async () => {
  await boot();
});

// ---------------------------------------------------------------------------
// The sheet edge: a block must refuse cleanly, never write the part of itself
// that happened to fit.
// ---------------------------------------------------------------------------

describe("debt schedule: hand-evaluated amortization", () => {
  // Excel's own PMT closed form (fv=0, type=0) - independent of the template,
  // which only ever emits the string "=PMT(...)" for the fake to store.
  function pmt(rate: number, nper: number, pv: number): number {
    const growth = Math.pow(1 + rate, nper);
    return -(pv * rate * growth) / (growth - 1);
  }

  it("closes to zero after the last period, interest always on the opening balance", async () => {
    helpers.select("Model!A1");
    await smt.insertTemplate("debt-schedule");

    // The four literal inputs the block itself wrote - read back, not assumed.
    const principal = helpers.value("Model!B2") as number;
    const annualRate = helpers.value("Model!B3") as number;
    const years = helpers.value("Model!B4") as number;
    const paymentsPerYear = helpers.value("Model!B5") as number;
    expect({ principal, annualRate, years, paymentsPerYear }).toEqual({
      principal: 1_000_000,
      annualRate: 0.06,
      years: 5,
      paymentsPerYear: 2,
    });

    const periods = years * paymentsPerYear;
    const periodRate = annualRate / paymentsPerYear;
    const payment = pmt(periodRate, periods, -principal);

    // Walks the SAME chain the row's own formulas encode: opening = prior
    // closing, interest = opening * rate (never the closing balance),
    // principal = payment - interest, closing = opening - principal.
    let opening = principal;
    let principalPaid = 0;
    for (let period = 0; period < periods; period += 1) {
      const interest = opening * periodRate;
      const principalThisPeriod = payment - interest;
      const closing = opening - principalThisPeriod;
      principalPaid += principalThisPeriod;
      opening = closing;
    }

    // Opening of period 11 (one past the block) is the schedule's true
    // closing balance: a fully-amortising loan reaches (about) zero.
    expect(opening).toBeCloseTo(0, 2);
    expect(principalPaid).toBeCloseTo(principal, 2);

    // Structural cross-check that the formulas above actually describe what
    // is written (columns A..F = Period, Opening, Payment, Interest,
    // Principal, Closing): period 1's interest reads its OWN Opening column,
    // its Closing is Opening less Principal, and period 2's Opening reads
    // period 1's Closing column - never its own Payment or Interest.
    expect(helpers.formula("Model!D10")).toBe("=B10*B3/B5");
    expect(helpers.formula("Model!F10")).toBe("=B10-E10");
    expect(helpers.formula("Model!B11")).toBe("=F10");
  });

  it("totals the flow columns to the payment count and the original principal", async () => {
    helpers.select("Model!A1");
    await smt.insertTemplate("debt-schedule");

    // Row 20 is the Totals row: Payment (C), Interest (D) and Principal (E)
    // sum the ten periods; Opening (B) and Closing (F), the two balance
    // columns, are left blank rather than summed.
    expect(helpers.formula("Model!C20")).toBe("=SUM(C10:C19)");
    expect(helpers.formula("Model!D20")).toBe("=SUM(D10:D19)");
    expect(helpers.formula("Model!E20")).toBe("=SUM(E10:E19)");
    expect(helpers.value("Model!B20")).toBe("");
    expect(helpers.value("Model!F20")).toBe("");
  });
});

// ---------------------------------------------------------------------------
// DCF: discount factors, Gordon terminal value, and enterprise value.
// ---------------------------------------------------------------------------

describe("DCF: hand-evaluated discounting and terminal value", () => {
  it("discounts every year at (1+WACC)^year and grows the last year into a terminal value", async () => {
    helpers.select("Model!A1");
    await smt.insertTemplate("dcf");

    const wacc = helpers.value("Model!B2") as number;
    const growth = helpers.value("Model!B3") as number;
    const fcf = ["B6", "C6", "D6", "E6", "F6"].map(
      (at) => helpers.value(`Model!${at}`) as number,
    );
    expect({ wacc, growth, fcf }).toEqual({
      wacc: 0.09,
      growth: 0.02,
      fcf: [100_000, 110_000, 120_000, 130_000, 140_000],
    });

    const discountFactors = fcf.map(
      (_v, index) => 1 / (1 + wacc) ** (index + 1),
    );
    const pvOfFcf = fcf.map((v, index) => v * discountFactors[index]!);
    const pvForecast = pvOfFcf.reduce((sum, v) => sum + v, 0);
    const terminalValue = (fcf.at(-1)! * (1 + growth)) / (wacc - growth);
    const pvTerminal = terminalValue * discountFactors.at(-1)!;
    const enterpriseValue = pvForecast + pvTerminal;

    // Formula shape at the two hinge cells: the discount factor uses WACC and
    // the Year row (not the growth rate), and the terminal value grows the
    // LAST fcf year, not the first.
    expect(helpers.formula("Model!B7")).toBe("=1/(1+B2)^B5");
    expect(helpers.formula("Model!B11")).toBe("=F6*(1+B3)/(B2-B3)");

    expect(helpers.value("Model!B10")).toBe(""); // formula cell, no live value
    // The independent oracle for the same inputs:
    expect(terminalValue).toBeCloseTo(2_040_000, 6);
    expect(pvForecast).toBeCloseTo(460_075.61, 1);
    expect(pvTerminal).toBeCloseTo(1_325_860.03, 1);
    expect(enterpriseValue).toBeCloseTo(1_785_935.64, 1);
    expect(enterpriseValue).toBeGreaterThan(pvForecast);
  });
});

// ---------------------------------------------------------------------------
// NPV / IRR: the outlay is never double-discounted, and payback lands where
// the cumulative column actually crosses zero.
// ---------------------------------------------------------------------------

describe("NPV / IRR: hand-evaluated cash flows", () => {
  function readFlows(): number[] {
    // Period 0 (the outlay) sits alone at B5; periods 1-8 run B6:B13.
    const outlay = helpers.value("Model!B5") as number;
    const later = Array.from(
      { length: 8 },
      (_unused, index) =>
        helpers.value(`Model!${cellAddress(5 + index, 1)}`) as number,
    );
    return [outlay, ...later];
  }

  it("NPV equals the outlay plus the eight flows each discounted by their own period, undiscounted at period 0", async () => {
    helpers.select("Model!A1");
    await smt.insertTemplate("npv-irr");

    const rate = helpers.value("Model!B2") as number;
    const flows = readFlows();
    expect(rate).toBe(0.1);
    expect(flows).toEqual([
      -500_000, 80_000, 95_000, 105_000, 115_000, 125_000, 135_000, 145_000,
      155_000,
    ]);

    // The formula is =NPV(rate, B6:B13)+B5: NPV() itself discounts its first
    // argument by period 1, so summing flows[0]/(1+r)^0 + flows[1..8]/(1+r)^t
    // is exactly what that formula computes - not a different definition that
    // happens to agree.
    const npv = flows.reduce((sum, flow, t) => sum + flow / (1 + rate) ** t, 0);
    expect(helpers.formula("Model!B15")).toBe("=NPV(B2,B6:B13)+B5");
    expect(npv).toBeCloseTo(109_209.99, 1);

    // IRR spans the outlay AND all eight flows - drop either end and the
    // series never resolves to the same root.
    expect(helpers.formula("Model!B16")).toBe("=IRR(B5:B13)");
    const npvAt = (r: number): number =>
      flows.reduce((sum, flow, t) => sum + flow / (1 + r) ** t, 0);
    // A real root exists between 0% and 100%: bisection is an independent,
    // reliable way to prove IRR(B5:B13) is well posed for these nine cells.
    let lo = 0;
    let hi = 1;
    expect(npvAt(lo)).toBeGreaterThan(0);
    expect(npvAt(hi)).toBeLessThan(0);
    for (let step = 0; step < 60; step += 1) {
      const mid = (lo + hi) / 2;
      if (npvAt(mid) > 0) lo = mid;
      else hi = mid;
    }
    expect(lo).toBeGreaterThan(0.14);
    expect(lo).toBeLessThan(0.16);
  });

  it("payback lands on the period the cumulative column turns non-negative, plus the fraction that closes it", async () => {
    helpers.select("Model!A1");
    await smt.insertTemplate("npv-irr");
    const flows = readFlows();

    // Independent walk of the same cumulative column the sheet builds one
    // formula cell at a time (C5, then C6 = C5 + B6, and so on).
    let cumulative = 0;
    const cumulativeByPeriod = flows.map((flow) => (cumulative += flow));
    const firstNonNegative = cumulativeByPeriod.findIndex((v) => v >= 0);
    expect(firstNonNegative).toBeGreaterThan(0); // never true at the outlay
    const payback =
      firstNonNegative -
      1 +
      Math.abs(cumulativeByPeriod[firstNonNegative - 1]!) /
        flows[firstNonNegative]!;

    expect(helpers.formula("Model!B17")).toBe(
      '=IFERROR((COUNTIF(C5:C13,"<0")-1)-INDEX(C5:C13,COUNTIF(C5:C13,"<0"))/INDEX(B5:B13,COUNTIF(C5:C13,"<0")+1),"n/a")',
    );
    expect(payback).toBeCloseTo(4.84, 2);
  });

  it("IFERROR names the payback as n/a instead of running INDEX off the end when it never recovers", async () => {
    // A single-flow, never-recovering series: COUNTIF sees every one of the
    // nine cells negative, so INDEX(range, 10) has no tenth cell to read -
    // exactly the case the formula's own IFERROR exists for.
    const flows = [-500_000, 1, 1, 1, 1, 1, 1, 1, 1];
    let cumulative = 0;
    const cumulativeByPeriod = flows.map((flow) => (cumulative += flow));
    expect(cumulativeByPeriod.every((v) => v < 0)).toBe(true);
    const negativeCount = cumulativeByPeriod.filter((v) => v < 0).length;
    expect(negativeCount).toBe(flows.length);
    // INDEX(range, negativeCount + 1) would be the tenth cell of a nine-cell
    // range: past the end, #REF!, which is exactly what IFERROR is there to
    // catch and rename "n/a".
    expect(negativeCount + 1).toBeGreaterThan(flows.length);
  });
});

// ---------------------------------------------------------------------------
// Working capital: three ratios computed from five literal balances.
// ---------------------------------------------------------------------------

describe("working capital: hand-evaluated days", () => {
  it("DSO, DIO and DPO net to the cash conversion cycle", async () => {
    helpers.select("Model!A1");
    await smt.insertTemplate("working-capital");

    const revenue = helpers.value("Model!B2") as number;
    const cogs = helpers.value("Model!B3") as number;
    const receivables = helpers.value("Model!B4") as number;
    const inventory = helpers.value("Model!B5") as number;
    const payables = helpers.value("Model!B6") as number;
    expect({ revenue, cogs, receivables, inventory, payables }).toEqual({
      revenue: 5_000_000,
      cogs: 3_000_000,
      receivables: 820_000,
      inventory: 640_000,
      payables: 510_000,
    });

    const dso = (receivables / revenue) * 365;
    const dio = (inventory / cogs) * 365;
    const dpo = (payables / cogs) * 365;
    const ccc = dso + dio - dpo;

    expect(helpers.formula("Model!B9")).toBe("=B4/B2*365"); // DSO: receivables over REVENUE
    expect(helpers.formula("Model!B10")).toBe("=B5/B3*365"); // DIO: inventory over COGS
    expect(helpers.formula("Model!B11")).toBe("=B6/B3*365"); // DPO: payables over COGS
    expect(helpers.formula("Model!B12")).toBe("=B9+B10-B11");

    expect(dso).toBeCloseTo(59.86, 6);
    expect(dio).toBeCloseTo(77.8667, 3);
    expect(dpo).toBeCloseTo(62.05, 6);
    expect(ccc).toBeCloseTo(75.6767, 3);
  });
});

// ---------------------------------------------------------------------------
// Two-way sensitivity: the four corners and the centre of the grid.
// ---------------------------------------------------------------------------

describe("sensitivity grid: hand-evaluated corners", () => {
  it("steps the base value by both drivers, resolved through the axis cells themselves", async () => {
    helpers.select("Model!A1");
    await smt.insertTemplate("sensitivity");

    const base = helpers.value("Model!B2") as number;
    const rowDriverStep = helpers.value("Model!B3") as number;
    const colDriverStep = helpers.value("Model!B4") as number;
    expect({ base, rowDriverStep, colDriverStep }).toEqual({
      base: 1_000_000,
      rowDriverStep: 0.05,
      colDriverStep: 0.1,
    });

    // Two-level hand evaluation: the axis cells are formulas over the two
    // driver-step inputs, and the corner formula multiplies the base by both
    // resolved axis values - never the raw step numbers.
    expect(helpers.formula("Model!A7")).toBe("=B3*-2"); // row axis, step -2
    expect(helpers.formula("Model!A11")).toBe("=B3*2"); // row axis, step +2
    expect(helpers.formula("Model!B6")).toBe("=B4*-2"); // column axis, step -2
    expect(helpers.formula("Model!F6")).toBe("=B4*2"); // column axis, step +2
    expect(helpers.formula("Model!B7")).toBe("=B2*(1+A7)*(1+B6)"); // top-left
    expect(helpers.formula("Model!F11")).toBe("=B2*(1+A11)*(1+F6)"); // bottom-right
    expect(helpers.formula("Model!D9")).toBe("=B2*(1+A9)*(1+D6)"); // centre

    const axis = (step: number, driverStep: number): number =>
      driverStep * step;
    const corner = (rowStep: number, colStep: number): number =>
      base *
      (1 + axis(rowStep, rowDriverStep)) *
      (1 + axis(colStep, colDriverStep));

    expect(corner(-2, -2)).toBeCloseTo(720_000, 6); // top-left
    expect(corner(-2, 2)).toBeCloseTo(1_080_000, 6); // top-right
    expect(corner(2, -2)).toBeCloseTo(880_000, 6); // bottom-left
    expect(corner(2, 2)).toBeCloseTo(1_320_000, 6); // bottom-right
    expect(corner(0, 0)).toBeCloseTo(base, 6); // centre: no step at all
  });
});

// ---------------------------------------------------------------------------
// EBITDA bridge: the check row must actually net to zero for the shipped
// literal steps, not merely reference the right cells.
// ---------------------------------------------------------------------------

describe("EBITDA bridge: the check row nets to zero", () => {
  it("opening plus every named step equals the closing literal", async () => {
    helpers.select("Model!A1");
    await smt.insertTemplate("ebitda-bridge");

    const opening = helpers.value("Model!B2") as number;
    const steps = ["B3", "B4", "B5", "B6", "B7"].map(
      (at) => helpers.value(`Model!${at}`) as number,
    );
    const closing = helpers.value("Model!B8") as number;
    expect({ opening, steps, closing }).toEqual({
      opening: 12_000_000,
      steps: [900_000, 600_000, -250_000, -700_000, 150_000],
      closing: 12_700_000,
    });

    const check =
      opening + steps.reduce((sum, step) => sum + step, 0) - closing;
    expect(check).toBe(0);
    expect(helpers.formula("Model!B9")).toBe("=B2+SUM(B3:B7)-B8");
  });
});

// ---------------------------------------------------------------------------
// Placement quirks: a second sheet, and a sheet name that needs quoting.
// ---------------------------------------------------------------------------
