// The details block under the PowerPoint pane's toast: the per-link lines an
// action stages while it runs (a row that failed, a source that changed, a
// chart that had to be a picture), which the guard reads once and drops.
// Owns that one buffer, so the pane holds no mutable state of its own for it.
// Invariant: what an action staged belongs to that action - the guard clears
// it in its finally, so nothing is ever shown beside the next one's message.

import { describeError, type ReportContext } from "../ui/report";

export interface PaneDetails {
  // What the toast shows under the message, or nothing.
  readonly value: string | undefined;
  set(lines: string | undefined): void;
  add(line: string): void;
  // A step that failed inside an action that still has something to report.
  addFailure(error: unknown, what: string): void;
}

export function createPaneDetails(context: ReportContext): PaneDetails {
  let value: string | undefined;
  function add(line: string): void {
    value = value === undefined ? line : `${value}\n${line}`;
  }
  return {
    get value(): string | undefined {
      return value;
    },
    set(lines) {
      value = lines;
    },
    add,
    addFailure(error, what) {
      add(`${what}: ${describeError(error, context).message}`);
    },
  };
}
