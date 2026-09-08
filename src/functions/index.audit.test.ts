// The independence rule of the custom-functions bundle: Office loads it into
// the function runtime, a page with no pane, no DOM and no Office.js object
// model. This suite proves both halves - the module graph reaches nothing but
// the pure allocator, and the functions evaluate with none of those globals.
import { readFileSync } from "node:fs";
import { dirname, join, relative } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";

const ENTRY = "src/functions/index.ts";
// What a pane module reaches for and this bundle may not. The header comments
// name Office.js and the DOM, so the comments come off before the scan.
const PANE_GLOBALS = [
  /\bdocument\b/,
  /\blocalStorage\b/,
  /\bwindow\b/,
  /\bOffice\./,
];

function code(file: string): string {
  return readFileSync(join(process.cwd(), file), "utf8").replace(
    /\/\/.*$/gm,
    "",
  );
}

class FakeFunctionError {
  constructor(
    public code: string,
    public message?: string,
  ) {}
}

// Every module the entry reaches, entry included, as repo-relative paths.
function moduleGraph(entry: string): string[] {
  const seen: string[] = [];
  const walk = (file: string): void => {
    if (seen.includes(file)) return;
    seen.push(file);
    const source = readFileSync(join(process.cwd(), file), "utf8");
    for (const match of source.matchAll(/from\s+"(\.[^"]*)"/g)) {
      const target = join(dirname(file), match[1] ?? "");
      walk(`${relative(process.cwd(), target)}.ts`);
    }
  };
  walk(entry);
  return seen;
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("the custom-functions bundle", () => {
  it("reaches the pure allocator and nothing else", () => {
    expect(moduleGraph(ENTRY)).toEqual([ENTRY, "src/rounding.ts"]);
  });

  it("names no pane global anywhere in that graph", () => {
    for (const file of moduleGraph(ENTRY)) {
      const source = code(file);
      for (const global of PANE_GLOBALS) {
        expect([file, global.test(source)]).toEqual([file, false]);
      }
    }
  });

  it("evaluates with no document and no Office", async () => {
    expect(typeof document).toBe("undefined");
    expect(typeof (globalThis as Record<string, unknown>).Office).toBe(
      "undefined",
    );

    const associated = new Map<string, unknown>();
    vi.stubGlobal("CustomFunctions", {
      associate: (id: string, implementation: unknown) => {
        associated.set(id, implementation);
      },
      Error: FakeFunctionError,
      ErrorCode: { invalidValue: "#VALUE!" },
    });
    vi.resetModules();
    const module = await import("./index");

    expect([...associated.keys()].sort()).toEqual(["ROUND", "ROUNDSUM"]);
    const group = [[1.005], [2.005], [3.005]];
    expect(module.smtRoundSum(group, 2)).toBe(6.02);
    expect([1, 2, 3].map((index) => module.smtRound(group, index, 2))).toEqual([
      1.01, 2.01, 3,
    ]);
  });
});
