// Golden test for the manifest renderer: proves buildManifest reproduces the
// committed manifest.prod.xml byte for byte, and that dev/prod differ only in
// the header comment and base URL. Also proves the top-level <Requirements>
// block is Workbook-only, every ribbon group and its label resource render
// the expected number of times, every ribbon FunctionName is registered in
// src/main.ts, and every interpolated value is XML-escaped.
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

  it("keeps the Workbook host at five ribbon groups (Office's per-tab cap is 6), Presentation at one", () => {
    const workbook = ADDIN.hosts.find((host) => host.name === "Workbook")!;
    const presentation = ADDIN.hosts.find(
      (host) => host.name === "Presentation",
    )!;
    expect(workbook.groups).toHaveLength(5);
    expect(workbook.groups.length).toBeLessThanOrEqual(6);
    expect(presentation.groups).toHaveLength(1);

    const xml = buildManifest(prod, ADDIN);
    const workbookBlocks = xml.match(
      /<Host xsi:type="Workbook">[\s\S]*?<\/Host>/g,
    )!;
    expect(workbookBlocks).toHaveLength(2);
    for (const block of workbookBlocks) {
      expect(block.match(/<Group id="SMT\.Group\./g)).toHaveLength(5);
    }
  });

  it("drops the top-level ExcelApi requirement once a second host exists, keeps SharedRuntime", () => {
    // Built from WORKBOOK_HOST directly, not ADDIN.hosts: ADDIN itself is
    // already two hosts, so this isolates the renderer's host-count behavior
    // from what the real spec currently declares.
    const singleHost: AddinSpec = { ...ADDIN, hosts: [WORKBOOK_HOST] };
    const withPpt: AddinSpec = {
      ...ADDIN,
      hosts: [
        WORKBOOK_HOST,
        {
          name: "Presentation",
          page: "pptpane.html",
          urlResid: "SMT.Pptpane.Url",
          taskpaneId: "SMT.Pptpane",
          groups: [
            { id: "SMT.Group.Links", label: "Model Tools Links", buttons: [] },
          ],
        },
      ],
    };
    const single = buildManifest(prod, singleHost);
    const multi = buildManifest(prod, withPpt);
    // Exactly-2-space indent targets the OfficeApp-level block only; the
    // VersionOverrides SharedRuntime blocks sit at 4/6-space indent and stay.
    expect(single).toMatch(/^ {2}<Requirements>$/m);
    expect(single).toContain(`<Set Name="ExcelApi"`);
    expect(multi).not.toMatch(/^ {2}<Requirements>$/m);
    expect(multi).not.toContain(`<Set Name="ExcelApi"`);
    expect(multi).toContain(`<bt:Set Name="SharedRuntime"`);
  });

  it("declares the Presentation host in both blocks and drops the top-level ExcelApi requirement", () => {
    const xml = buildManifest(prod, ADDIN);
    expect(xml.match(/<Host xsi:type="Presentation">/g)).toHaveLength(2);
    expect(xml).toContain('<Host Name="Presentation"/>');
    expect(xml).not.toContain('<Set Name="ExcelApi"');
    expect(xml).toContain(
      '<bt:Url id="SMT.Pptpane.Url" DefaultValue="https://dbautomatizacijas.com/modelis/pptpane.html"/>',
    );
  });

  it("wires custom functions into the Workbook host only, once per block", () => {
    const xml = buildManifest(prod, ADDIN);
    // Once in each VersionOverrides block, and never on the PowerPoint host.
    expect(
      xml.match(/<ExtensionPoint xsi:type="CustomFunctions">/g),
    ).toHaveLength(2);
    for (const block of xml.match(
      /<Host xsi:type="Presentation">[\s\S]*?<\/Host>/g,
    )!) {
      expect(block).not.toContain("CustomFunctions");
      expect(block).not.toContain("<AllFormFactors>");
    }

    const workbookBlocks = xml.match(
      /<Host xsi:type="Workbook">[\s\S]*?<\/Host>/g,
    )!;
    expect(workbookBlocks).toHaveLength(2);
    for (const block of workbookBlocks) {
      expect(
        block.match(/<ExtensionPoint xsi:type="CustomFunctions">/g),
      ).toHaveLength(1);
      // The schema fixes the order inside <Host>, and the functions reuse the
      // pane's runtime rather than declaring a second one.
      expect(block.indexOf("<Runtimes>")).toBeLessThan(
        block.indexOf("<AllFormFactors>"),
      );
      expect(block.indexOf("<AllFormFactors>")).toBeLessThan(
        block.indexOf("<DesktopFormFactor>"),
      );
      expect(block.match(/<Runtime resid=/g)).toHaveLength(1);
      // Indentation differs between the two blocks, so the point is compared
      // with its whitespace collapsed.
      const point = /<AllFormFactors>[\s\S]*?<\/AllFormFactors>/
        .exec(block)![0]
        .replaceAll(/\s+/g, " ");
      expect(point).toContain(
        '<Script> <SourceLocation resid="SMT.Functions.Script.Url"/> </Script>',
      );
      // The page is the pane the shared runtime already serves.
      expect(point).toContain(
        '<Page> <SourceLocation resid="SMT.Taskpane.Url"/> </Page>',
      );
      expect(point).toContain(
        '<Metadata> <SourceLocation resid="SMT.Functions.Metadata.Url"/> </Metadata>',
      );
      expect(point).toContain('<Namespace resid="SMT.Functions.Namespace"/>');
    }
  });

  it("publishes the functions script, metadata and namespace as resources", () => {
    const xml = buildManifest(prod, ADDIN);
    for (const line of [
      '<bt:Url id="SMT.Functions.Script.Url" DefaultValue="https://dbautomatizacijas.com/modelis/functions.js"/>',
      '<bt:Url id="SMT.Functions.Metadata.Url" DefaultValue="https://dbautomatizacijas.com/modelis/functions.json"/>',
      '<bt:String id="SMT.Functions.Namespace" DefaultValue="SMT"/>',
    ]) {
      // One per <Resources> block, and there are two.
      expect(xml.split(line)).toHaveLength(3);
    }
    // A host without the block renders neither the extension point nor its ids.
    const noFunctions: AddinSpec = {
      ...ADDIN,
      hosts: [{ ...WORKBOOK_HOST, customFunctions: undefined }],
    };
    const plain = buildManifest(prod, noFunctions);
    expect(plain).not.toContain("CustomFunctions");
    expect(plain).not.toContain("SMT.Functions.Script.Url");
  });

  it("emits each group label resource exactly once per VersionOverrides block", () => {
    const xml = buildManifest(prod, ADDIN);
    const resourceBlocks = [
      ...xml.matchAll(/<Resources>[\s\S]*?<\/Resources>/g),
    ].map((match) => match[0]);
    expect(resourceBlocks).toHaveLength(2);

    const groupIds = ADDIN.hosts.flatMap((host) =>
      host.groups.map((group) => group.id),
    );
    expect(groupIds.length).toBeGreaterThan(0);
    for (const block of resourceBlocks) {
      for (const groupId of groupIds) {
        const pattern = new RegExp(
          `<bt:String id="${groupId.replaceAll(".", "\\.")}\\.Label"`,
          "g",
        );
        expect(block.match(pattern)).toHaveLength(1);
      }
    }
  });

  // The generated-bytes gate compares the build to the committed file, so it
  // cannot see a duplicate id: both sides would carry it. Only the renderer can.
  it("refuses a spec whose two hosts reuse one button id", () => {
    const pptWithOpenPane: AddinSpec = {
      ...ADDIN,
      hosts: [
        WORKBOOK_HOST,
        {
          name: "Presentation",
          page: "pptpane.html",
          urlResid: "SMT.Pptpane.Url",
          taskpaneId: "SMT.Pptpane",
          groups: [
            {
              id: "SMT.Group.Links",
              label: "Model Tools Links",
              // The obvious next edit: the same pane button on both hosts.
              buttons: [
                {
                  id: "OpenPane",
                  label: "Links",
                  tip: "Open the Model Tools linked-objects pane.",
                  action: { kind: "showPane" },
                },
              ],
            },
          ],
        },
      ],
    };
    expect(() => buildManifest(prod, pptWithOpenPane)).toThrow(
      "manifest: duplicate resource id SMT.OpenPane.Label",
    );
  });

  it("refuses a spec whose two hosts reuse one group id", () => {
    const sharedGroup: AddinSpec = {
      ...ADDIN,
      hosts: [
        WORKBOOK_HOST,
        {
          name: "Presentation",
          page: "pptpane.html",
          urlResid: "SMT.Pptpane.Url",
          taskpaneId: "SMT.Pptpane",
          groups: [
            { id: "SMT.Group.Tools", label: "Model Tools", buttons: [] },
          ],
        },
      ],
    };
    expect(() => buildManifest(prod, sharedGroup)).toThrow(
      "manifest: duplicate resource id SMT.Group.Tools.Label",
    );
  });

  it("every ribbon FunctionName is registered in src/main.ts registerCommands", () => {
    const xml = buildManifest(prod, ADDIN);
    const functionNames = new Set(
      [...xml.matchAll(/<FunctionName>([A-Z_]+)<\/FunctionName>/g)].map(
        (match) => match[1]!,
      ),
    );
    expect(functionNames.size).toBeGreaterThan(0);

    const mainSrc = readFileSync(
      new URL("../src/main.ts", import.meta.url),
      "utf8",
    );
    const registered = new Set(
      [...mainSrc.matchAll(/(SMT_[A-Z_]+):/g)].map((match) => match[1]!),
    );
    for (const name of functionNames) {
      expect(registered.has(name)).toBe(true);
    }
  });

  it('escapes & and " in interpolated text', () => {
    const spec: AddinSpec = {
      ...ADDIN,
      hosts: [
        {
          ...WORKBOOK_HOST,
          groups: [
            {
              id: "SMT.Group.Tools",
              label: "Model Tools",
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
          groups: [
            {
              id: "SMT.Group.Tools",
              label: "Model Tools",
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
        },
      ],
    };
    const xml = buildManifest(prod, spec);
    expect(xml).toContain("&lt;script&gt;alert(&apos;x&apos;)&lt;/script&gt;");
    expect(xml).not.toContain("<script>");
  });
});
