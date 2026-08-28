// Golden test for the manifest renderer: proves buildManifest reproduces the
// committed manifest.prod.xml byte for byte, and that dev/prod differ only in
// the header comment and base URL. Also proves every host renders in both
// VersionOverrides blocks (V1_0 outer, V1_1 nested).
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { ADDIN, ENVIRONMENTS } from "./spec";
import { buildManifest } from "./xml";

const prod = ENVIRONMENTS.find((env) => env.name === "prod")!;
const dev = ENVIRONMENTS.find((env) => env.name === "dev")!;

describe("buildManifest", () => {
  it("reproduces the committed production manifest byte for byte", () => {
    const committed = readFileSync(new URL("../manifest.prod.xml", import.meta.url), "utf8");
    expect(buildManifest(prod, ADDIN)).toBe(committed);
  });

  it("dev differs from prod only in the header comment and the base URL", () => {
    const normalise = (xml: string) =>
      xml
        .replace(/<!--[\s\S]*?-->/, "")
        .replaceAll("https://localhost:3000/", "BASE/")
        .replaceAll("https://dbautomatizacijas.com/modelis/", "BASE/");
    expect(normalise(buildManifest(dev, ADDIN))).toBe(normalise(buildManifest(prod, ADDIN)));
  });

  it("emits every host in both VersionOverrides blocks", () => {
    const xml = buildManifest(prod, ADDIN);
    expect(xml.match(/<Host xsi:type="Workbook">/g)).toHaveLength(2);
    expect(xml.match(/<Host Name="Workbook"\/>/g)).toHaveLength(1);
  });
});
