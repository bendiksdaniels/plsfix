// The debouncer auto-push waits on: one flush per quiet window, keys that fall
// due together handed over together, and nothing left to fire after a cancel.

import { afterEach, describe, expect, it, vi } from "vitest";
import { virtualClock } from "../../test/clock";
import { Debouncer, systemClock } from "./debounce";

const DELAY = 3000;

function collector(): { calls: string[][]; flush: (keys: string[]) => void } {
  const calls: string[][] = [];
  return {
    calls,
    flush: (keys) => {
      calls.push([...keys].sort());
    },
  };
}

describe("Debouncer", () => {
  it("waits out the delay before flushing", () => {
    const clock = virtualClock();
    const sink = collector();
    const debouncer = new Debouncer(DELAY, sink.flush, clock);

    debouncer.touch(["Model"]);
    clock.advance(DELAY - 1);
    expect(sink.calls).toEqual([]);

    clock.advance(1);
    expect(sink.calls).toEqual([["Model"]]);
    expect(debouncer.pending()).toEqual([]);
  });

  // Typing is a burst of edits, not one: the window has to start again at each
  // keystroke or every one of them would push.
  it("defers the flush while the same key keeps being touched", () => {
    const clock = virtualClock();
    const sink = collector();
    const debouncer = new Debouncer(DELAY, sink.flush, clock);

    debouncer.touch(["Model"]);
    clock.advance(1000);
    debouncer.touch(["Model"]);
    clock.advance(2500);
    expect(sink.calls).toEqual([]);

    clock.advance(500);
    expect(sink.calls).toEqual([["Model"]]);
  });

  it("hands over every key that fell due in one call", () => {
    const clock = virtualClock();
    const sink = collector();
    const debouncer = new Debouncer(DELAY, sink.flush, clock);

    debouncer.touch(["Model"]);
    debouncer.touch(["Data"]);
    clock.advance(DELAY);
    expect(sink.calls).toEqual([["Data", "Model"]]);
  });

  // Each key has its own deadline, so a sheet still being typed on cannot hold
  // back a sheet that went quiet.
  it("flushes a settled key while a later one keeps waiting", () => {
    const clock = virtualClock();
    const sink = collector();
    const debouncer = new Debouncer(DELAY, sink.flush, clock);

    debouncer.touch(["Model"]);
    clock.advance(1000);
    debouncer.touch(["Data"]);
    clock.advance(2000);
    expect(sink.calls).toEqual([["Model"]]);
    expect(debouncer.pending()).toEqual(["Data"]);

    clock.advance(1000);
    expect(sink.calls).toEqual([["Model"], ["Data"]]);
  });

  it("drops everything waiting when it is cancelled", () => {
    const clock = virtualClock();
    const sink = collector();
    const debouncer = new Debouncer(DELAY, sink.flush, clock);

    debouncer.touch(["Model"]);
    debouncer.cancel();
    clock.advance(DELAY * 2);
    expect(sink.calls).toEqual([]);
    expect(clock.waiting()).toBe(0);
  });
});

describe("systemClock", () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it("schedules and cancels on the host timer", () => {
    vi.useFakeTimers();
    const clock = systemClock();
    const ran: string[] = [];

    clock.after(10, () => ran.push("kept"));
    clock.after(10, () => ran.push("dropped")).cancel();
    vi.advanceTimersByTime(10);

    expect(ran).toEqual(["kept"]);
    expect(clock.now()).toBe(Date.now());
  });
});
