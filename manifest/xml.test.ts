// Golden test for the manifest renderer: proves buildManifest reproduces the
// committed manifest.prod.xml byte for byte, and that dev/prod differ only in
// the header comment and base URL. Also proves the top-level <Requirements>
// block is Workbook-only and every interpolated value is XML-escaped.
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { ADDIN, ENVIRONMENTS, WORKBOOK_HOST } from "./spec";
import type { AddinSpec } from "./spec";
import { buildManifest } from "./xml";

const prod = ENVIRONMENTS.find((env) => env.name === "prod")!;
const dev = ENVIRONMENTS.find((env) => env.name === "dev")!;

describe("buildManifest", () => {
  it("reproduces the committed production manifest byte for byte", () => {
    const committed = readFileSync(
      new URL("../manifest.prod.xml", import.meta.url),
      "utf8",
    );
    expect(buildManifest(prod, ADDIN)).toBe(committed);
  });

  it("dev differs from prod only in the header comment and the base URL", () => {
    const normalise = (xml: string) =>
      xml
        .replace(/<!--[\s\S]*?-->/, "")
        .replaceAll("https://localhost:3000/", "BASE/")
        .replaceAll("https://dbautomatizacijas.com/modelis/", "BASE/");
    expect(normalise(buildManifest(dev, ADDIN))).toBe(
      normalise(buildManifest(prod, ADDIN)),
    );
  });

  it("emits every host in both VersionOverrides blocks", () => {
    const xml = buildManifest(prod, ADDIN);
    expect(xml.match(/<Host xsi:type="Workbook">/g)).toHaveLength(2);
    expect(xml.match(/<Host Name="Workbook"\/>/g)).toHaveLength(1);
  });

  it("drops the top-level ExcelApi requirement once a second host exists, keeps SharedRuntime", () => {
    const withPpt: AddinSpec = {
      ...ADDIN,
      hosts: [
        ...ADDIN.hosts,
        {
          name: "Presentation",
          page: "pptpane.html",
          urlResid: "SMT.Pptpane.Url",
          taskpaneId: "SMT.Pptpane",
          groupId: "SMT.Group.Links",
          groupLabel: "Model Tools Links",
          buttons: [],
        },
      ],
    };
    const single = buildManifest(prod, ADDIN);
    const multi = buildManifest(prod, withPpt);
    // Exactly-2-space indent targets the OfficeApp-level block only; the
    // VersionOverrides SharedRuntime blocks sit at 4/6-space indent and stay.
    expect(single).toMatch(/^ {2}<Requirements>$/m);
    expect(single).toContain(`<Set Name="ExcelApi"`);
    expect(multi).not.toMatch(/^ {2}<Requirements>$/m);
    expect(multi).not.toContain(`<Set Name="ExcelApi"`);
    expect(multi).toContain(`<bt:Set Name="SharedRuntime"`);
  });

  it('escapes & and " in interpolated text', () => {
    const spec: AddinSpec = {
      ...ADDIN,
      hosts: [
        {
          ...WORKBOOK_HOST,
          buttons: [
            {
              id: "OpenPane",
              label: "Model Tools",
              tip: 'Fill & go "now"',
              action: { kind: "showPane" },
            },
          ],
        },
      ],
    };
    const xml = buildManifest(prod, spec);
    expect(xml).toContain("&amp;");
    expect(xml).toContain("&quot;");
    expect(xml).not.toContain('Fill & go "now"');
  });

  it("escapes <, > and ' in interpolated text", () => {
    const spec: AddinSpec = {
      ...ADDIN,
      hosts: [
        {
          ...WORKBOOK_HOST,
          buttons: [
            {
              id: "OpenPane",
              label: "Model Tools",
              tip: "<script>alert('x')</script>",
              action: { kind: "showPane" },
            },
          ],
        },
      ],
    };
    const xml = buildManifest(prod, spec);
    expect(xml).toContain("&lt;script&gt;alert(&apos;x&apos;)&lt;/script&gt;");
    expect(xml).not.toContain("<script>");
  });
});
