// Chart maths kept out of the Office.js layer so it is testable on its own.

import { type CellValue } from "./model";

// Where a series really starts and ends. A modeller selects the whole row a
// growth line sits on, blank cells at its ends included; those state nothing, so
// they are not periods. Null when the line holds nothing at all.
export function seriesSpan(
  cells: CellValue[],
): { first: number; last: number } | null {
  const blank = (value: CellValue | undefined): boolean =>
    value === undefined || value === null || value === "";
  let first = 0;
  let last = cells.length - 1;
  while (first <= last && blank(cells[first])) first += 1;
  while (last >= first && blank(cells[last])) last -= 1;
  return first > last ? null : { first, last };
}

export interface BridgeSeries {
  base: number[];
  rise: number[];
  fall: number[];
}

const MIN_BRIDGE_POINTS = 3;

// A bridge reads first-total, deltas, last-total. Stacking base under rise and
// fall composes the classic waterfall; base + rise is the running level after
// every point, which is what reconciles the deltas against the closing total.
export function bridgeSeries(values: number[]): BridgeSeries {
  if (values.length < MIN_BRIDGE_POINTS) {
    throw new Error(
      "A bridge needs an opening total, a delta and a closing total.",
    );
  }

  const base: number[] = [];
  const rise: number[] = [];
  const fall: number[] = [];
  let level = 0;

  values.forEach((value, index) => {
    if (index === 0 || index === values.length - 1) {
      base.push(0);
      rise.push(value);
      fall.push(0);
      level = value;
      return;
    }
    if (value < 0) {
      // Falls are plotted as a magnitude sitting on top of where they land.
      base.push(level + value);
      rise.push(0);
      fall.push(-value);
      level += value;
      return;
    }
    base.push(level);
    rise.push(value);
    fall.push(0);
    level += value;
  });

  return { base, rise, fall };
}

export function cagr(first: number, last: number, periods: number): number {
  if (periods < 1) throw new Error("A CAGR needs at least one period.");
  if (first <= 0 || last <= 0) {
    throw new Error("A CAGR needs a positive start and end value.");
  }
  const ratio = last / first;
  if (Number.isFinite(ratio)) return ratio ** (1 / periods) - 1;
  // The ratio overflowed a double although a rate may still exist (1e-300 to
  // 1e300 over 100 periods is 999 999): take the root in log space, and let a
  // result that is still infinite say so to the caller.
  return Math.exp((Math.log(last) - Math.log(first)) / periods) - 1;
}

export function formatCagrLabel(value: number): string {
  // Rounding before the sign test keeps a hair below zero from printing "-0.0%".
  const percent = Math.round(value * 1000) / 10;
  const sign = percent < 0 ? "-" : "+";
  return `CAGR ${sign}${Math.abs(percent).toFixed(1)}%`;
}

export interface TornadoDriver {
  label: string;
  low: number;
  high: number;
}

export interface TornadoSeries {
  labels: string[];
  low: number[];
  high: number[];
  base: number;
}

const MIN_TORNADO_DRIVERS = 2;

// A tornado ranks drivers by how far they move the answer, widest swing first.
// Both series are deltas from the base, so one bar row per driver spans its low
// and high around the base line. A null base means the model never stated one:
// the mean of every low and high is then the neutral middle of the range the
// drivers describe.
export function tornadoSeries(
  drivers: TornadoDriver[],
  base: number | null,
): TornadoSeries {
  if (drivers.length < MIN_TORNADO_DRIVERS) {
    throw new Error("tornado: need at least two drivers");
  }
  const outcomes = drivers.flatMap((driver) => [driver.low, driver.high]);
  if (!outcomes.every((value) => Number.isFinite(value))) {
    throw new Error("tornado: every low and high must be a finite number");
  }
  if (base !== null && !Number.isFinite(base)) {
    throw new Error("tornado: the base must be a finite number");
  }

  const middle =
    base ?? outcomes.reduce((sum, value) => sum + value, 0) / outcomes.length;
  // A stable sort, so drivers that move the answer equally keep the order they
  // were selected in instead of swapping places between runs.
  const ranked = [...drivers].sort(
    (left, right) =>
      Math.abs(right.high - right.low) - Math.abs(left.high - left.low),
  );

  return {
    labels: ranked.map((driver) => driver.label),
    low: ranked.map((driver) => driver.low - middle),
    high: ranked.map((driver) => driver.high - middle),
    base: middle,
  };
}

export interface FootballRow {
  label: string;
  low: number;
  high: number;
}

export interface FootballField {
  labels: string[];
  low: number[];
  high: number[];
  /** high - low: the band a stacked bar draws on top of the invisible floor. */
  range: number[];
  /** Which rows arrived with their low above their high and were swapped. */
  swapped: boolean[];
  swaps: number;
}

const MIN_FOOTBALL_ROWS = 2;

// A football field stacks an invisible bar up to each method's low and a
// visible one across its range, so the floating band is the valuation. The
// rows keep the order they were selected in: a football field is read top down
// in the order the banker laid the methods out, not by width.
export function footballField(rows: FootballRow[]): FootballField {
  if (rows.length < MIN_FOOTBALL_ROWS) {
    throw new Error("football field: need at least two rows");
  }
  const bounds = rows.flatMap((row) => [row.low, row.high]);
  if (!bounds.every((value) => Number.isFinite(value))) {
    throw new Error(
      "football field: every low and high must be a finite number",
    );
  }

  const ordered = rows.map((row) => ({
    label: row.label,
    low: Math.min(row.low, row.high),
    high: Math.max(row.low, row.high),
    swapped: row.low > row.high,
  }));

  return {
    labels: ordered.map((row) => row.label),
    low: ordered.map((row) => row.low),
    high: ordered.map((row) => row.high),
    range: ordered.map((row) => row.high - row.low),
    swapped: ordered.map((row) => row.swapped),
    swaps: ordered.filter((row) => row.swapped).length,
  };
}
