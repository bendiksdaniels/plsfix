import { defineConfig } from "vitest/config";
import pkg from "./package.json" with { type: "json" };

export default defineConfig({
  // Same __APP_VERSION__ global as vite.config.ts, so tests see the real value.
  define: {
    __APP_VERSION__: JSON.stringify(pkg.version),
    __PUBLIC_VERSION__: JSON.stringify(pkg.publicVersion ?? pkg.version),
  },
  test: {
    // The clock-driven "host never answers" suites (test/j.*, ppt.charts.audit)
    // step fake timers ten times per test; on a loaded GitHub runner that
    // crossed vitest's 5 s default (12.09, four timeouts in one run) while a
    // hung test still fails, only later.
    testTimeout: 20_000,
    include: [
      "src/**/*.test.ts",
      "test/**/*.test.ts",
      "manifest/**/*.test.ts",
      "scripts/**/*.test.ts",
    ],
    // On since the hunt-order pass (27.09.2026): proven order-independent
    // across 10 whole-suite shuffled seeds after fixing the five files that
    // were not (three vi.mock isolation leaks, two click-through suites
    // whose "coverage" tests read what earlier tests pressed, now pinned
    // per-suite with `{ shuffle: false }`). A run under this prints its own
    // seed first ("Running tests with seed \"<n>\"", verified on vitest
    // 4.1.11); reproduce a red run with `--sequence.seed=<n>`.
    sequence: { shuffle: true },
  },
});
