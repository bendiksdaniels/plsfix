import { defineConfig } from "vitest/config";
import pkg from "./package.json" with { type: "json" };

export default defineConfig({
  // Same __APP_VERSION__ global as vite.config.ts, so tests see the real value.
  define: { __APP_VERSION__: JSON.stringify(pkg.version) },
  test: {
    include: ["src/**/*.test.ts", "test/**/*.test.ts", "manifest/**/*.test.ts"],
  },
});
