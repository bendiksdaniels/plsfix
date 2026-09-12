// The rig the comps stress suites share: a strict fake Excel host, the real
// pane dispatch over the real adapters (only src/pane/shared.ts, which wants a
// DOM, is mocked away) and the assertions every adversarial scenario reuses.
//
// Owns: the boot order. Invariant: dispatch() and ../src/excel are the same
// module instances, so a flow driven through the pane path writes into the
// workbook these helpers read back.

import { expect, vi } from "vitest";
import {
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

/**
 * What src/pane/shared.ts is replaced by: every export the pane modules pull
 * off it, none of which a headless stress run needs. Only isExcelReady is real
 * behaviour - dispatch refuses every action while it answers false.
 */
export function paneShared(): Record<string, unknown> {
  return {
    APP_VERSION: "v0.0.000",
    DELETE_CONFIRM_MS: 0,
    actionButtons: [],
    errorMessage: (error: unknown) => String(error),
    guard: async (run: () => Promise<string>) => {
      await run();
    },
    isExcelReady: () => true,
    refreshSelection: async () => undefined,
    renderActionState: () => undefined,
    setExcelReady: () => undefined,
    tabs: { activate: () => undefined },
    toast: { show: () => undefined },
  };
}

export interface Rig {
  helpers: FakeHelpers;
  workbook: FakeWorkbook;
  smt: typeof ExcelModule;
  /** The real src/pane/dispatch.ts, over the real adapters. */
  dispatch(action: string): Promise<string>;
}

export async function boot(sheets = ["Model", "Data"]): Promise<Rig> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets });
  const smt = await import("../src/excel");
  const { dispatch } = await import("../src/pane/dispatch");
  return { helpers: host.helpers, workbook: host.workbook, smt, dispatch };
}

/** The line the pane would show: the result, or the refusal's own sentence. */
export async function sentence(run: () => Promise<string>): Promise<string> {
  try {
    return await run();
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
}

/**
 * Every line these tools answer with - a result or a refusal - opens with the
 * tool's own name, so office.js's own string ("The worksheet 'Model' is
 * protected...") fails here instead of reaching the toast.
 */
export function expectSentence(stage: string, line: string): void {
  expect(
    line.startsWith(stage),
    `not a ${stage} sentence: ${JSON.stringify(line)}`,
  ).toBe(true);
}

/** A three-column comps table with a header row at A1:C4. */
export function seedComps(rig: Rig): void {
  rig.helpers.seed("Model!A1", [
    ["Company", "EV/EBITDA", "Listed"],
    ["Alpha", 8.1, "yes"],
    ["Beta", 9.4, "yes"],
    ["Gamma", 7.2, "no"],
  ]);
  rig.helpers.select("Model!A1:C4");
}

/** Three valuation methods with a header row at A1:C4. */
export function seedMethods(rig: Rig): void {
  rig.helpers.seed("Model!A1", [
    ["Method", "Low", "High"],
    ["DCF", 90, 130],
    ["Trading comps", 100, 120],
    ["Precedents", 110, 150],
  ]);
  rig.helpers.select("Model!A1:C4");
}

/** A five-row block of numbers at A1:C5. */
export function seedGrid(rig: Rig): void {
  rig.helpers.seed(
    "Model!A1",
    Array.from({ length: 5 }, (_unused, row) => [row, row + 1, row + 2]),
  );
  rig.helpers.select("Model!A1:C5");
}

/** Ctrl+A: the whole grid, which Excel counts as -1 cells. */
export const WHOLE_SHEET = "Model!A1:XFD1048576";
/** 5 000 cells, the selection cap exactly; 5 001 is one past it. */
export const AT_CAP = "Model!A1:E1000";
export const PAST_CAP = "Model!A1:C1667";
export const UNDER_CAP = "Model!A1:E999";
