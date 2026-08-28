// Chart maths kept out of the Office.js layer so it is testable on its own.

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
  return (last / first) ** (1 / periods) - 1;
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
