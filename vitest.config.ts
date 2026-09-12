import { defineConfig } from "vitest/config";
import pkg from "./package.json" with { type: "json" };

export default defineConfig({
  // Same __APP_VERSION__ global as vite.config.ts, so tests see the real value.
  define: { __APP_VERSION__: JSON.stringify(pkg.version) },
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
  },
});
