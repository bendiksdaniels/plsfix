// The footer shows the public release line (package.json publicVersion) while
// error reports keep the real build (package.json version): both reach the
// panes through vite's defines, and both footers read the public one.
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import pkg from "../../package.json" with { type: "json" };
import { formatVersion } from "./version";

describe("the two versions a pane knows", () => {
  it("defines the public line and the real build from package.json", () => {
    expect(__PUBLIC_VERSION__).toBe(pkg.publicVersion);
    expect(__APP_VERSION__).toBe(pkg.version);
    expect(formatVersion(__PUBLIC_VERSION__)).toMatch(/^v\d+\.\d+\.\d{3}$/);
  });

  it("puts the public line in both footers, never the build", () => {
    for (const file of ["src/main.ts", "src/ppt/main.ts"]) {
      expect(readFileSync(file, "utf8")).toContain(
        'getElement("app-version").textContent = DISPLAY_VERSION;',
      );
    }
  });
});
