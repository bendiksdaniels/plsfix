# Tier 1 Groundwork Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the codebase ready for a second Office host: one manifest source, one version source, a split Excel layer, shared pane UI, an error surface, and a single `npm run check` gate that CI runs.

**Architecture:** Pure mechanical split of `src/excel.ts` behind a barrel; a TypeScript manifest spec that generates both manifests; version injected by Vite from `package.json`; shared UI modules under `src/ui/` extracted from `src/main.ts`; eslint + prettier + a GitHub Actions workflow.

**Tech Stack:** TypeScript 7 (strict), Vite 8, Vitest 4, tsx (new dev dep, runs TS scripts), eslint 9 + typescript-eslint, prettier 3, jsdom (ui tests), Rust/axum server unchanged.

**Spec:** `docs/superpowers/specs/2026-08-28-ppt-links-design.md` (section 9 "Prerequisites").

## Global Constraints

- Existing 219 vitest tests and 6 cargo tests stay green after every task; no test is edited to pass (test edits only for moved import paths).
- No `any`, no `@ts-ignore`, no `console.*`, no TODO/FIXME in `src/`, `test/`, `scripts/`, `manifest/`, `server/`.
- `README.md` and `ROADMAP.md` are hand-curated: surgical edits only, never regenerate.
- Commit messages: terse, lower-case, Daniel's voice, NO Co-Authored-By trailer (e.g. `split excel.ts into src/excel/`).
- Node 22+ (Daniel runs Node 25). Dev manifest sideload is Mac-only (`office-addin-debugging`).
- Daniel's Excel `wef` file `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/FF1B34D8-DD7D-4B39-8FA9-6248CA09DB6E.manifest.xml` is a HARD LINK to the repo's `manifest.xml` today. Task 3 replaces it with a plain copy of `manifest.prod.xml` BEFORE any task rewrites `manifest.xml`. Generators must write files via a temp file + rename (never truncate in place).
- Execution order: Task 1 alone first; Tasks 2, 4, 6 in parallel worktrees; Task 5 last (it reformats every file).

---

### Task 1: Split `src/excel.ts` into `src/excel/*` behind a barrel

**Files:**
- Create: `src/excel/index.ts`, `src/excel/shared.ts`, `src/excel/undo.ts`, `src/excel/selection.ts`, `src/excel/formulas.ts`, `src/excel/paste.ts`, `src/excel/autocolor.ts`, `src/excel/audit.ts`, `src/excel/trace.ts`, `src/excel/charts.ts`, `src/excel/workbook.ts`
- Delete: `src/excel.ts`
- Test: unchanged `test/host.integration.test.ts` (imports `../src/excel`, which now resolves to the directory index)

**Interfaces:**
- Consumes: the 49 current exports of `src/excel.ts` (list them with `grep -n "^export" src/excel.ts`).
- Produces: `src/excel/index.ts` re-exporting exactly the same names and types, so `import { ... } from "./excel"` in `src/main.ts` and `../src/excel` in tests are untouched.

Section boundaries in today's file (line numbers from `grep -n "^// ----" src/excel.ts`):

| New file | Lines of `src/excel.ts` | Content |
|---|---|---|
| `shared.ts` | 1-135 (public part) | `PresetName`, `NumberFormatName`, `SelectionSummary`, `parseAddress` (the names the old file exported) |
| `internal.ts` | 1-135 (private part) + line 1189 | `SELECTION_CELL_CAP`, range helpers (`writeRuns`, local address helpers), fill helpers, `hostSupports`; imported by siblings and by `src/excel/links.ts` later, never re-exported by the barrel |
| `undo.ts` | 136-245 | SMT Undo: `captureUndo`, `undoTarget`, `lastUndoSkipped`, `undoLastAction` |
| `selection.ts` | 246-519 | `selectionWithinCap`, `inspectSelection`, presets, number formats, cycles, `clearFormats` |
| `formulas.ts` | 520-676 | fast fill, IFERROR, scaling, sign flip, decimals, CAGR |
| `paste.ts` | 677-787 | copy source + paste special + paste exact |
| `autocolor.ts` | 788-1000 | autocolor + color key + on-edit handler |
| `audit.ts` | 1001-1166 | overlay snapshot/restore/toggle (`persistOverlaySetting`, `restorePersistedOverlay`, `snapshotFills`, `restoreFills`, `toggleAuditOverlay`) |
| `trace.ts` | 1167-1252 | Smart Track (`traceActiveCell`, `selectArea`) |
| `charts.ts` | 1253-1503 | chart shell, waterfall, `formatSelectedChart`, `addCagrLabel` |
| `workbook.ts` | 1504-1694 | TOC, sheet explorer, name scrubber |

Module-private helpers used across sections move to `internal.ts` and become exported there (internal API of the folder; the barrel does NOT re-export them). Exact names come from the compiler: everything a sibling file cannot see after the move.

- [ ] **Step 1: Record the export list and the baseline**

```bash
cd ~/plsfix
grep -n "^export" src/excel.ts | sed 's/(.*//' > /tmp/excel-exports-before.txt
wc -l < /tmp/excel-exports-before.txt   # expect 49
npm test 2>&1 | tail -5                  # expect 219 passed
```

- [ ] **Step 2: Create the folder and move each section verbatim**

Use `sed -n 'A,Bp' src/excel.ts > src/excel/<file>.ts` per row of the table, then add the imports each file needs (the compiler tells you: run `npx tsc --noEmit` and fix every "Cannot find name" by importing from `./shared`, `./undo`, etc.). Keep function bodies byte-identical; only imports/exports change.

- [ ] **Step 3: Write the barrel**

```ts
// src/excel/index.ts
// The Office.js layer, one file per feature area. The pane and the tests import
// this barrel, so the split is invisible to them.
export * from "./shared";
export * from "./undo";
export * from "./selection";
export * from "./formulas";
export * from "./paste";
export * from "./autocolor";
export * from "./audit";
export * from "./trace";
export * from "./charts";
export * from "./workbook";
```

`shared.ts` holds only the four public names; everything else the sections share lives in `internal.ts`, which the barrel does not export.

- [ ] **Step 4: Delete the old file and prove the export set is identical**

```bash
git rm -q src/excel.ts
npx tsc --noEmit
node -e '
const fs=require("fs");
const names=new Set();
for (const f of fs.readdirSync("src/excel")) {
  if (f==="index.ts"||f==="internal.ts") continue;
  for (const m of fs.readFileSync("src/excel/"+f,"utf8").matchAll(/^export (?:async function|function|const|type|interface) (\w+)/gm)) names.add(m[1]);
}
const before=fs.readFileSync("/tmp/excel-exports-before.txt","utf8").match(/(?:function|const|type|interface) (\w+)/g).map(s=>s.split(" ")[1]);
const missing=before.filter(n=>!names.has(n));
if (missing.length) { console.error("MISSING", missing); process.exit(1); }
console.log("exports intact:", before.length);
'
```
Expected: `exports intact: 49`.

- [ ] **Step 5: Run the whole suite**

Run: `npm test`
Expected: 219 passed (same count as Step 1). `npm run build` green.

- [ ] **Step 6: Commit**

```bash
git add -A src/excel
git commit -m "split excel.ts into src/excel/ behind a barrel"
```

---

### Task 2: Manifest generator (one spec, two files, both hosts later)

**Files:**
- Create: `manifest/spec.ts`, `manifest/xml.ts`, `manifest/xml.test.ts`, `scripts/build-manifests.ts`
- Modify: `package.json` (scripts + `tsx` dev dep), `manifest.xml` (regenerated as the DEV manifest), `manifest.prod.xml` (regenerated, must be byte-identical to today's file)
- Test: `manifest/xml.test.ts`

**Interfaces:**
- Produces: `buildManifest(env: ManifestEnvironment, spec: AddinSpec): string` (pure, returns the XML text), `ENVIRONMENTS`, `ADDIN` (spec data), `HOSTS: HostSpec[]`. Task 8 of the feature plan adds a `Presentation` entry to `HOSTS`; nothing else changes.
- Scripts: `npm run manifest:build` writes both files; `npm run manifest:check` exits 1 when a committed file differs from its generated text.

- [ ] **Step 1: Add tsx and the scripts**

```bash
npm install --save-dev tsx@4
```
`package.json` scripts (add; keep the others):
```json
"manifest:build": "tsx scripts/build-manifests.ts",
"manifest:check": "tsx scripts/build-manifests.ts --check",
"validate": "office-addin-manifest validate manifest.xml && office-addin-manifest validate manifest.prod.xml"
```

- [ ] **Step 2: Write the spec data (exact values from today's manifest)**

```ts
// manifest/spec.ts
// Single source of truth for both manifests. Everything host- or URL-specific
// lives here; manifest/xml.ts only renders it.
import { readFileSync } from "node:fs";

export interface ManifestEnvironment {
  name: "dev" | "prod";
  file: string;
  baseUrl: string;
  comment: string;
}

export interface ButtonSpec {
  id: string; // e.g. "Autocolor" -> control id SMT.Button.Autocolor
  label: string;
  tip: string;
  action: { kind: "showPane" } | { kind: "function"; name: string };
}

export interface HostSpec {
  name: "Workbook" | "Presentation";
  page: string; // pane html served from baseUrl
  urlResid: string; // resource id of the pane URL
  taskpaneId: string;
  groupId: string;
  groupLabel: string;
  buttons: ButtonSpec[];
}

export interface AddinSpec {
  id: string;
  version: string; // "X.Y.Z.0", derived from package.json
  provider: string;
  displayName: string;
  description: string;
  supportUrl: string;
  appDomain: string;
  tabLabel: string;
  iconResids: { 16: string; 32: string; 80: string };
  hosts: HostSpec[];
}

const packageVersion = (
  JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")) as {
    version: string;
  }
).version;

export const ENVIRONMENTS: readonly ManifestEnvironment[] = [
  {
    name: "dev",
    file: "manifest.xml",
    baseUrl: "https://localhost:3000/",
    comment:
      "Development manifest: same GUID as manifest.prod.xml (same add-in identity), only URLs differ.\n" +
      "     Sideloaded by npm start; the pane is served by vite on https://localhost:3000/.\n" +
      "     GENERATED by scripts/build-manifests.ts from manifest/spec.ts - do not edit by hand.",
  },
  {
    name: "prod",
    file: "manifest.prod.xml",
    baseUrl: "https://dbautomatizacijas.com/modelis/",
    comment:
      "Production manifest: same GUID as manifest.xml (same add-in identity), only URLs differ.\n" +
      "     Served LIVE from https://dbautomatizacijas.com/modelis/ (suite key modelis, Rust host :8804).\n" +
      "     The path carries a Cloudflare Access bypass - Office webviews cannot pass Access\n" +
      "     (docs/research/launch-path.md), so only built pane assets may live there.",
  },
];

export const WORKBOOK_HOST: HostSpec = {
  name: "Workbook",
  page: "taskpane.html",
  urlResid: "SMT.Taskpane.Url",
  taskpaneId: "SMT.Taskpane",
  groupId: "SMT.Group.Tools",
  groupLabel: "Model Tools",
  buttons: [
    { id: "OpenPane", label: "Model Tools", tip: "Open the Model Tools task pane.", action: { kind: "showPane" } },
    { id: "Autocolor", label: "Autocolor", tip: "Color inputs, formulas and links in your palette.", action: { kind: "function", name: "SMT_AUTOCOLOR" } },
    { id: "FillRight", label: "Fill Right", tip: "Fill the formula from the left cell across the selection.", action: { kind: "function", name: "SMT_FILLRIGHT" } },
    { id: "FillDown", label: "Fill Down", tip: "Fill the formula from the top cell down the selection.", action: { kind: "function", name: "SMT_FILLDOWN" } },
    { id: "IfError", label: "IFERROR", tip: "Wrap selected formulas with IFERROR.", action: { kind: "function", name: "SMT_IFERROR" } },
  ],
};

export const ADDIN: AddinSpec = {
  id: "FF1B34D8-DD7D-4B39-8FA9-6248CA09DB6E",
  version: `${packageVersion}.0`,
  provider: "Daniels Bendiks",
  displayName: "Model Tools",
  description: "Fast, consistent financial modelling tools for Excel.",
  supportUrl: "https://dbautomatizacijas.com",
  appDomain: "https://dbautomatizacijas.com",
  tabLabel: "Model Tools",
  iconResids: { 16: "SMT.Icon.16", 32: "SMT.Icon.32", 80: "SMT.Icon.80" },
  hosts: [WORKBOOK_HOST],
};
```

Note: today's manifest says `<Version>1.0.0.0</Version>` while package.json is 1.1.0. The generated prod file must be byte-identical to the committed one in Step 5, so for this task ONLY, hard-code `version: "1.0.0.0"` with a comment `// Task 4 switches this to package.json`; Task 4 replaces it with the derived value and bumps the committed manifests in the same commit.

- [ ] **Step 3: Write the failing renderer test**

```ts
// manifest/xml.test.ts
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
```

Run: `npx vitest run manifest/xml.test.ts` → FAIL (module `./xml` missing). Add `"manifest/**/*.test.ts"` to `vitest.config.ts` `include`.

- [ ] **Step 4: Write the renderer**

Structure the emitter so both VersionOverrides blocks come from ONE function called with two indents. Whitespace must match the committed file exactly: 2-space indentation, `xsi:type` attribute order as in the file, self-closing tags without a space before `/>`.

```ts
// manifest/xml.ts
import type { AddinSpec, ButtonSpec, HostSpec, ManifestEnvironment } from "./spec";

const pad = (block: string, spaces: number): string =>
  block
    .split("\n")
    .map((line) => (line ? " ".repeat(spaces) + line : line))
    .join("\n");

function icons(spec: AddinSpec, indent: number): string {
  return pad(
    [
      "<Icon>",
      `  <bt:Image size="16" resid="${spec.iconResids[16]}"/>`,
      `  <bt:Image size="32" resid="${spec.iconResids[32]}"/>`,
      `  <bt:Image size="80" resid="${spec.iconResids[80]}"/>`,
      "</Icon>",
    ].join("\n"),
    indent,
  );
}

function control(spec: AddinSpec, host: HostSpec, button: ButtonSpec): string {
  const action =
    button.action.kind === "showPane"
      ? [
          `<Action xsi:type="ShowTaskpane">`,
          `  <TaskpaneId>${host.taskpaneId}</TaskpaneId>`,
          `  <SourceLocation resid="${host.urlResid}"/>`,
          `</Action>`,
        ]
      : [
          `<Action xsi:type="ExecuteFunction">`,
          `  <FunctionName>${button.action.name}</FunctionName>`,
          `</Action>`,
        ];
  return [
    `<Control xsi:type="Button" id="SMT.Button.${button.id}">`,
    `  <Label resid="SMT.${button.id}.Label"/>`,
    `  <Supertip>`,
    `    <Title resid="SMT.${button.id}.Label"/>`,
    `    <Description resid="SMT.${button.id}.Tip"/>`,
    `  </Supertip>`,
    icons(spec, 2),
    ...action.map((line) => `  ${line}`),
    `</Control>`,
  ].join("\n");
}

function hostBlock(spec: AddinSpec, host: HostSpec): string {
  return [
    `<Host xsi:type="${host.name}">`,
    `  <Runtimes>`,
    `    <Runtime resid="${host.urlResid}" lifetime="long"/>`,
    `  </Runtimes>`,
    `  <DesktopFormFactor>`,
    `    <FunctionFile resid="${host.urlResid}"/>`,
    `    <ExtensionPoint xsi:type="PrimaryCommandSurface">`,
    `      <CustomTab id="SMT.Tab">`,
    `        <Group id="${host.groupId}">`,
    `          <Label resid="${host.groupId}.Label"/>`,
    pad(icons(spec, 0), 10),
    ...host.buttons.map((button) => pad(control(spec, host, button), 10)),
    `        </Group>`,
    `        <Label resid="SMT.Tab.Label"/>`,
    `      </CustomTab>`,
    `    </ExtensionPoint>`,
    `  </DesktopFormFactor>`,
    `</Host>`,
  ].join("\n");
}

function resources(env: ManifestEnvironment, spec: AddinSpec): string {
  const shorts = [
    `<bt:String id="SMT.Tab.Label" DefaultValue="${spec.tabLabel}"/>`,
    ...spec.hosts.flatMap((host) => [
      `<bt:String id="${host.groupId}.Label" DefaultValue="${host.groupLabel}"/>`,
      ...host.buttons.map((b) => `<bt:String id="SMT.${b.id}.Label" DefaultValue="${b.label}"/>`),
    ]),
  ];
  const longs = spec.hosts.flatMap((host) =>
    host.buttons.map((b) => `<bt:String id="SMT.${b.id}.Tip" DefaultValue="${b.tip}"/>`),
  );
  return [
    `<Resources>`,
    `  <bt:Images>`,
    `    <bt:Image id="${spec.iconResids[16]}" DefaultValue="${env.baseUrl}assets/icon-16.png"/>`,
    `    <bt:Image id="${spec.iconResids[32]}" DefaultValue="${env.baseUrl}assets/icon-32.png"/>`,
    `    <bt:Image id="${spec.iconResids[80]}" DefaultValue="${env.baseUrl}assets/icon-80.png"/>`,
    `  </bt:Images>`,
    `  <bt:Urls>`,
    ...spec.hosts.map((host) => `    <bt:Url id="${host.urlResid}" DefaultValue="${env.baseUrl}${host.page}"/>`),
    `  </bt:Urls>`,
    `  <bt:ShortStrings>`,
    ...shorts.map((line) => `    ${line}`),
    `  </bt:ShortStrings>`,
    `  <bt:LongStrings>`,
    ...longs.map((line) => `    ${line}`),
    `  </bt:LongStrings>`,
    `</Resources>`,
  ].join("\n");
}

function overrides(env: ManifestEnvironment, spec: AddinSpec, nested: string | null): string {
  const xmlns = nested === null
    ? `xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides/1.1" xsi:type="VersionOverridesV1_1"`
    : `xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0"`;
  return [
    `<VersionOverrides ${xmlns}>`,
    `  <Requirements>`,
    `    <bt:Sets DefaultMinVersion="1.1">`,
    `      <bt:Set Name="SharedRuntime" MinVersion="1.1"/>`,
    `    </bt:Sets>`,
    `  </Requirements>`,
    `  <Hosts>`,
    ...spec.hosts.map((host) => pad(hostBlock(spec, host), 4)),
    `  </Hosts>`,
    pad(resources(env, spec), 2),
    ...(nested === null
      ? [`  <ExtendedOverrides Url="${env.baseUrl}shortcuts.json"/>`]
      : [pad(nested, 2)]),
    `</VersionOverrides>`,
  ].join("\n");
}

export function buildManifest(env: ManifestEnvironment, spec: AddinSpec): string {
  const inner = overrides(env, spec, null);
  const outer = overrides(env, spec, inner);
  const primary = spec.hosts[0]!;
  return [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
    `<!-- ${env.comment} -->`,
    `<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"`,
    `  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"`,
    `  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"`,
    `  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"`,
    `  xsi:type="TaskPaneApp">`,
    `  <Id>${spec.id}</Id>`,
    `  <Version>${spec.version}</Version>`,
    `  <ProviderName>${spec.provider}</ProviderName>`,
    `  <DefaultLocale>en-US</DefaultLocale>`,
    `  <DisplayName DefaultValue="${spec.displayName}"/>`,
    `  <Description DefaultValue="${spec.description}"/>`,
    `  <IconUrl DefaultValue="${env.baseUrl}assets/icon-32.png"/>`,
    `  <SupportUrl DefaultValue="${spec.supportUrl}"/>`,
    `  <AppDomains>`,
    `    <AppDomain>${spec.appDomain}</AppDomain>`,
    `  </AppDomains>`,
    `  <Hosts>`,
    ...spec.hosts.map((host) => `    <Host Name="${host.name}"/>`),
    `  </Hosts>`,
    // A top-level requirement applies to every host; ExcelApi would hide the
    // add-in in PowerPoint, so it is declared only while Excel is the sole host.
    ...(spec.hosts.every((host) => host.name === "Workbook")
      ? [
          `  <Requirements>`,
          `    <Sets DefaultMinVersion="1.9">`,
          `      <Set Name="ExcelApi" MinVersion="1.9"/>`,
          `    </Sets>`,
          `  </Requirements>`,
        ]
      : []),
    `  <DefaultSettings>`,
    `    <SourceLocation DefaultValue="${env.baseUrl}${primary.page}"/>`,
    `  </DefaultSettings>`,
    `  <Permissions>ReadWriteDocument</Permissions>`,
    pad(outer, 2),
    `</OfficeApp>`,
    ``,
  ].join("\n");
}
```

The conditional `<Requirements>` block is what lets the feature plan add the Presentation host as pure data.

Iterate on the whitespace until the first test passes: `diff <(npx tsx -e 'import {ADDIN,ENVIRONMENTS} from "./manifest/spec"; import {buildManifest} from "./manifest/xml"; process.stdout.write(buildManifest(ENVIRONMENTS[1],ADDIN))') manifest.prod.xml` must print nothing.

- [ ] **Step 5: Write the build/check script**

```ts
// scripts/build-manifests.ts
// Writes manifest.xml (dev) and manifest.prod.xml (prod) from manifest/spec.ts.
// --check: exit 1 if a committed file differs (CI gate). Files are written via
// temp + rename so a hard-linked copy elsewhere never sees a half-written file.
import { readFileSync, renameSync, writeFileSync } from "node:fs";
import { ADDIN, ENVIRONMENTS } from "../manifest/spec";
import { buildManifest } from "../manifest/xml";

const check = process.argv.includes("--check");
let drift = 0;
for (const env of ENVIRONMENTS) {
  const expected = buildManifest(env, ADDIN);
  const path = new URL(`../${env.file}`, import.meta.url);
  let current = "";
  try {
    current = readFileSync(path, "utf8");
  } catch {
    current = "";
  }
  if (current === expected) continue;
  if (check) {
    process.stderr.write(`${env.file} differs from manifest/spec.ts - run npm run manifest:build\n`);
    drift += 1;
    continue;
  }
  const tmp = new URL(`../${env.file}.tmp`, import.meta.url);
  writeFileSync(tmp, expected);
  renameSync(tmp, path);
  process.stdout.write(`wrote ${env.file}\n`);
}
process.exit(drift ? 1 : 0);
```

- [ ] **Step 6: Generate, validate, run tests**

```bash
npm run manifest:build      # writes manifest.xml (dev URLs); manifest.prod.xml unchanged
git diff --stat             # only manifest.xml changed
npm run manifest:check      # exit 0
npm run validate            # both manifests valid
npm test                    # 219 + 3 new
```

- [ ] **Step 7: Commit**

```bash
git add manifest scripts/build-manifests.ts manifest.xml package.json package-lock.json vitest.config.ts
git commit -m "manifests generated from one spec; dev manifest back on localhost"
```

---

### Task 3: Dev loop that cannot break the production sideload

**Files:**
- Create: `scripts/wef-restore-prod.sh`
- Modify: `package.json` (`poststop`, `start`), `CLAUDE.md` (manifest lines), `README.md:100-125` (surgical: the two sentences describing `manifest.xml` as localhost are now true again; add the `npm stop` note)
- Test: manual (commands below)

**Interfaces:** none (shell + docs).

- [ ] **Step 1: Replace the hard link with a plain copy (one-time, do it FIRST, before Task 2 merges)**

```bash
WEF=~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
F=$WEF/FF1B34D8-DD7D-4B39-8FA9-6248CA09DB6E.manifest.xml
stat -f '%l %i' "$F"                     # link count 2 today
rm "$F" && cp ~/plsfix/manifest.prod.xml "$F"
stat -f '%l %i' "$F"                     # link count 1, new inode
grep -c "dbautomatizacijas.com/modelis" "$F"   # > 0: Excel still loads from the server
```
Do NOT run this while Daniel's Excel visual pass is in progress; ask if unsure.

- [ ] **Step 2: Restore script**

```bash
#!/bin/sh
# After `npm stop` (which unregisters the dev manifest) put the production
# manifest back into every Office wef folder that exists, so a fresh Excel or
# PowerPoint launch loads the add-in from the server again. Mac only.
set -eu
[ "$(uname)" = Darwin ] || exit 0
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
ID="FF1B34D8-DD7D-4B39-8FA9-6248CA09DB6E"
for app in com.microsoft.Excel com.microsoft.Powerpoint; do
  wef="$HOME/Library/Containers/$app/Data/Documents/wef"
  [ -d "$wef" ] || continue
  cp "$ROOT/manifest.prod.xml" "$wef/$ID.manifest.xml"
  echo "restored prod manifest for $app"
done
```
`chmod +x scripts/wef-restore-prod.sh`. package.json: `"poststop": "sh scripts/wef-restore-prod.sh"`.

- [ ] **Step 3: Verify the loop**

```bash
npm start -- --no-sideload   # dev server up, no Excel launch
npm stop                     # prints "restored prod manifest for com.microsoft.Excel"
grep -c modelis ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/*.manifest.xml
```

- [ ] **Step 4: Docs**

`CLAUDE.md`: replace the manifest bullet with: "`manifest.xml` (dev, localhost:3000) and `manifest.prod.xml` are GENERATED from `manifest/spec.ts` (`npm run manifest:build`, `manifest:check` in CI); never edit them by hand. `npm stop` restores the prod manifest into the wef folders (`scripts/wef-restore-prod.sh`)." `README.md`: surgical edit of the sideload paragraph only.

- [ ] **Step 5: Commit**

```bash
git add scripts/wef-restore-prod.sh package.json CLAUDE.md README.md
git commit -m "dev loop: poststop restores the prod manifest into wef"
```

---

### Task 4: One version source

**Files:**
- Create: `src/ui/version.ts`, `src/ui/version.test.ts`, `scripts/check-version.ts`
- Modify: `package.json` (scripts), `vite.config.ts` + `vitest.config.ts` (`define`), `src/vite-env.d.ts` (declare `__APP_VERSION__`), `taskpane.html:325-331` (footer), `src/main.ts` (fill the footer), `manifest/spec.ts` (version from package.json), `manifest.xml` + `manifest.prod.xml` (regenerated), `server/Cargo.toml` (stays 1.1.0 for now; the check only asserts equality)

**Interfaces:**
- Produces: `formatVersion(semver: string): string` (`"1.1.0"` -> `"v1.1.000"`), global `__APP_VERSION__: string`, `npm run version:check`.

- [ ] **Step 1: Failing test**

```ts
// src/ui/version.test.ts
import { describe, expect, it } from "vitest";
import { formatVersion } from "./version";

describe("formatVersion", () => {
  it("shows a three-digit patch", () => {
    expect(formatVersion("1.1.0")).toBe("v1.1.000");
    expect(formatVersion("2.0.12")).toBe("v2.0.012");
  });
  it("rejects anything that is not MAJOR.MINOR.PATCH", () => {
    expect(() => formatVersion("2.0")).toThrow(/semver/);
  });
});
```

- [ ] **Step 2: Implementation**

```ts
// src/ui/version.ts
// Footer format: vMAJOR.MINOR.PATCH with a three-digit patch (v1.1.002).
export function formatVersion(semver: string): string {
  const match = /^(\d+)\.(\d+)\.(\d+)$/.exec(semver);
  if (!match) throw new Error(`not a semver version: ${semver}`);
  return `v${match[1]}.${match[2]}.${match[3]!.padStart(3, "0")}`;
}
```
`vite.config.ts` and `vitest.config.ts`: `define: { __APP_VERSION__: JSON.stringify(pkg.version) }` with `import pkg from "./package.json" with { type: "json" }` (or `createRequire`). `src/vite-env.d.ts`: `declare const __APP_VERSION__: string;`. `taskpane.html` footer: replace the literal `v1.1.000` with `<span id="app-version"></span>`; `src/main.ts` at boot: `getElement("app-version").textContent = formatVersion(__APP_VERSION__);`.

- [ ] **Step 3: Version check script**

```ts
// scripts/check-version.ts
// package.json is the only place a version is typed by hand. Fails when the
// server crate, the manifests or the footer wiring disagree.
import { readFileSync } from "node:fs";
const root = new URL("../", import.meta.url);
const read = (file: string) => readFileSync(new URL(file, root), "utf8");
const version = (JSON.parse(read("package.json")) as { version: string }).version;
const problems: string[] = [];
if (!read("server/Cargo.toml").includes(`version = "${version}"`)) problems.push(`server/Cargo.toml is not ${version}`);
for (const file of ["manifest.xml", "manifest.prod.xml"]) {
  if (!read(file).includes(`<Version>${version}.0</Version>`)) problems.push(`${file} is not ${version}.0`);
}
if (read("taskpane.html").match(/v\d+\.\d+\.\d{3}/)) problems.push("taskpane.html hard-codes a version; use #app-version");
if (problems.length) { process.stderr.write(problems.join("\n") + "\n"); process.exit(1); }
process.stdout.write(`version ${version} consistent\n`);
```
package.json: `"version:check": "tsx scripts/check-version.ts"`. `manifest/spec.ts`: `version: \`${packageVersion}.0\`` (remove the Task 2 hard-code), `npm run manifest:build` (manifests now say 1.1.0.0).

- [ ] **Step 4: Verify**

`npm test` (+2), `npm run version:check` prints `version 1.1.0 consistent`, `npm run build` green, `grep -c "1.1.0.0" manifest.prod.xml` = 1.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "one version source: package.json feeds footer, manifests and the check"
```

---

### Task 5: `npm run check` + CI (runs LAST, reformats everything)

**Files:**
- Create: `eslint.config.js`, `.prettierrc.json`, `.prettierignore`, `.github/workflows/check.yml`
- Modify: `package.json` (dev deps + scripts), `tsconfig.json` (`include` adds `test`, `scripts`, `manifest`)

**Interfaces:** `npm run check` = typecheck + lint + format:check + test + test:server + manifest:check + version:check.

- [ ] **Step 1: Install**

```bash
npm install --save-dev eslint@9 typescript-eslint@8 prettier@3
```

- [ ] **Step 2: Configs**

```js
// eslint.config.js
import tseslint from "typescript-eslint";
export default tseslint.config(
  { ignores: ["dist/**", "node_modules/**", "server/**", "docs/**"] },
  ...tseslint.configs.strict,
  ...tseslint.configs.stylistic,
  {
    files: ["**/*.ts"],
    rules: {
      "no-console": "error",
      "@typescript-eslint/no-explicit-any": "error",
      "@typescript-eslint/no-non-null-assertion": "off",
      "@typescript-eslint/consistent-type-imports": "error",
    },
  },
);
```
`.prettierrc.json`: `{ "printWidth": 80, "trailingComma": "all" }` (matches the existing style). `.prettierignore`: `dist`, `node_modules`, `server/target`, `package-lock.json`, `*.md`, `manifest*.xml`.

package.json scripts:
```json
"typecheck": "tsc --noEmit",
"lint": "eslint .",
"format": "prettier --write .",
"format:check": "prettier --check .",
"test:server": "cargo test --manifest-path server/Cargo.toml",
"check": "npm run typecheck && npm run lint && npm run format:check && npm test && npm run test:server && npm run manifest:check && npm run version:check"
```
`tsconfig.json` include: `["src", "test", "scripts", "manifest", "vite.config.ts", "vitest.config.ts"]`; fix every error tsc now reports in `test/` (they are real; do not loosen the config).

- [ ] **Step 3: Format once, lint clean**

```bash
npm run format          # one formatting-only diff
npm run lint            # fix findings; no eslint-disable comments
npm run check           # all green
```

- [ ] **Step 4: Workflow**

```yaml
# .github/workflows/check.yml
name: check
on: [push, pull_request]
jobs:
  check:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with: { node-version: 22, cache: npm }
      - uses: dtolnay/rust-toolchain@stable
      - run: npm ci
      - run: npm run check
```

- [ ] **Step 5: Commit (two commits: format-only, then tooling)**

```bash
git add -A && git commit -m "prettier pass"
```
(then after the configs) `git commit -m "npm run check: tsc, eslint, prettier, tests, cargo, manifest and version gates; CI"`

---

### Task 6: Shared pane UI + error surface

**Files:**
- Create: `src/ui/toast.ts`, `src/ui/toast.test.ts`, `src/ui/guard.ts`, `src/ui/guard.test.ts`, `src/ui/tabs.ts`, `src/ui/tabs.test.ts`, `src/ui/report.ts`, `src/ui/report.test.ts`
- Modify: `src/main.ts` (use the modules; delete the local `showToast`, `setBusy` stays, `guard` becomes `makeGuard`, tab switcher replaced), `taskpane.html:332` (toast markup), `src/styles.css` (toast button)

**Interfaces (Produces, used verbatim by the PowerPoint pane later):**
```ts
// src/ui/toast.ts
export type ToastKind = "success" | "error";
export interface Toast { show(message: string, kind?: ToastKind, details?: string): void }
export function createToast(container: HTMLElement, hideAfterMs?: number): Toast;
// src/ui/guard.ts
export interface GuardDeps {
  setBusy(busy: boolean): void;
  notify(message: string, kind?: ToastKind, details?: string): void;
  describe(error: unknown, action: string | undefined): { message: string; details: string };
  after?(): Promise<void>;         // e.g. refreshSelection
  decorate?(message: string): string;
  finally?(): void;                // e.g. renderActionState
}
export type Guard = (run: () => Promise<string>, action?: string) => Promise<void>;
export function makeGuard(deps: GuardDeps): Guard;
// src/ui/tabs.ts
export function installTabs(bar: HTMLElement): { activate(tabId: string): void };
// src/ui/report.ts
export interface ReportContext { host: string; version: string }
export function describeError(error: unknown, ctx: ReportContext, action?: string): { message: string; details: string };
export function installErrorReporting(ctx: ReportContext, notify: (message: string, details: string) => void): void;
```

- [ ] **Step 1: Failing tests (jsdom)**

`npm install --save-dev jsdom@26` (Task 5 keeps it). Tests that touch the DOM start with the `// @vitest-environment jsdom` comment.

```ts
// src/ui/toast.test.ts
// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import { createToast } from "./toast";

describe("toast", () => {
  it("shows the message, then hides", () => {
    vi.useFakeTimers();
    const root = document.createElement("div");
    const toast = createToast(root, 100);
    toast.show("Saved");
    expect(root.className).toContain("visible");
    expect(root.textContent).toContain("Saved");
    vi.advanceTimersByTime(101);
    expect(root.className).toBe("toast");
  });
  it("offers a copy button only when details exist", async () => {
    const root = document.createElement("div");
    const write = vi.fn(() => Promise.resolve());
    Object.assign(navigator, { clipboard: { writeText: write } });
    createToast(root, 100).show("Failed", "error", "stack...");
    const button = root.querySelector("button")!;
    button.click();
    expect(write).toHaveBeenCalledWith("stack...");
    createToast(root, 100).show("Plain");
    expect(root.querySelector("button")).toBeNull();
  });
});
```
```ts
// src/ui/guard.test.ts
import { describe, expect, it, vi } from "vitest";
import { makeGuard } from "./guard";

describe("makeGuard", () => {
  it("busy on, run, after, decorate, notify, finally, busy off", async () => {
    const calls: string[] = [];
    const guard = makeGuard({
      setBusy: (b) => calls.push(`busy:${b}`),
      notify: (m, k) => calls.push(`notify:${k ?? "success"}:${m}`),
      describe: (e) => ({ message: String(e), details: "" }),
      after: async () => { calls.push("after"); },
      decorate: (m) => `${m}!`,
      finally: () => calls.push("finally"),
    });
    await guard(async () => "Done");
    expect(calls).toEqual(["busy:true", "after", "notify:success:Done!", "finally", "busy:false"]);
  });
  it("routes a failure to notify(error) with details and still clears busy", async () => {
    const notify = vi.fn();
    const guard = makeGuard({
      setBusy: () => undefined,
      notify,
      describe: (e, action) => ({ message: (e as Error).message, details: `action=${action}` }),
    });
    await guard(async () => { throw new Error("boom"); }, "export");
    expect(notify).toHaveBeenCalledWith("boom", "error", "action=export");
  });
});
```
```ts
// src/ui/tabs.test.ts
// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { installTabs } from "./tabs";

describe("installTabs", () => {
  it("shows exactly one panel and marks its tab selected", () => {
    document.body.innerHTML = `
      <nav id="bar"><button role="tab" id="t1" aria-controls="p1" aria-selected="true" class="tab active"></button>
      <button role="tab" id="t2" aria-controls="p2" aria-selected="false" class="tab"></button></nav>
      <div id="p1"></div><div id="p2" hidden></div>`;
    const tabs = installTabs(document.getElementById("bar")!);
    tabs.activate("t2");
    expect(document.getElementById("p1")!.hidden).toBe(true);
    expect(document.getElementById("p2")!.hidden).toBe(false);
    expect(document.getElementById("t2")!.getAttribute("aria-selected")).toBe("true");
    expect(document.getElementById("t1")!.classList.contains("active")).toBe(false);
  });
});
```
```ts
// src/ui/report.test.ts
import { describe, expect, it, vi } from "vitest";
import { describeError, installErrorReporting } from "./report";

const ctx = { host: "Excel", version: "v1.1.000" };
describe("describeError", () => {
  it("keeps the user message short and puts everything else in details", () => {
    const { message, details } = describeError(new Error("Select a chart first."), ctx, "export-chart");
    expect(message).toBe("Select a chart first.");
    expect(details).toContain("action: export-chart");
    expect(details).toContain("host: Excel");
    expect(details).toContain("version: v1.1.000");
    expect(details).toContain("Error: Select a chart first.");
  });
  it("names Office.js error codes and debug info when present", () => {
    const error = Object.assign(new Error("x"), { code: "ItemNotFound", debugInfo: { errorLocation: "Range.load" } });
    expect(describeError(error, ctx).details).toContain("code: ItemNotFound");
    expect(describeError(error, ctx).details).toContain("Range.load");
  });
  it("falls back for non-Error throws", () => {
    expect(describeError("nope", ctx).message).toBe("The add-in could not complete that action.");
  });
});
describe("installErrorReporting", () => {
  it("forwards window errors and unhandled rejections", () => {
    const notify = vi.fn();
    const listeners: Record<string, (event: unknown) => void> = {};
    const fakeWindow = { addEventListener: (type: string, fn: (event: unknown) => void) => { listeners[type] = fn; } };
    installErrorReporting(ctx, notify, fakeWindow as unknown as Window);
    listeners["error"]!({ error: new Error("late") });
    listeners["unhandledrejection"]!({ reason: new Error("async") });
    expect(notify).toHaveBeenCalledTimes(2);
    expect(notify.mock.calls[1]![0]).toBe("async");
  });
});
```
Run: `npx vitest run src/ui` → FAIL (modules missing).

- [ ] **Step 2: Implement the four modules**

`toast.ts`: keeps the existing behaviour (`className = "toast visible <kind>"`, hide after 3200 ms) and renders `<span class="toast-text">` + optional `<button class="toast-copy" type="button">Copy details</button>` whose click calls `navigator.clipboard.writeText(details)` and falls back to selecting a hidden `<textarea>` when the clipboard API is unavailable (Office webviews may deny it).

`guard.ts`: exactly the sequence the test pins; `describe` supplies message + details; the error path calls `notify(message, "error", details)`.

`tabs.ts`: reads `[role=tab]` buttons under `bar`, each with `aria-controls`; clicking or `activate(id)` sets `hidden` on every panel except the target, toggles `active` class and `aria-selected`. Also clears no children (the audit pane's stale-listener lesson is about re-render, handled by the callers).

`report.ts`: `describeError` builds `details` lines: `message`, `action: ...`, `host: ...`, `version: ...`, `platform: navigator.userAgent`, `code: ...` and `debugInfo: JSON` when present on the error, then the stack. `installErrorReporting(ctx, notify, win = window)` registers `error` and `unhandledrejection` listeners that call `notify(message, details)`.

- [ ] **Step 3: Wire `src/main.ts`**

Replace the local `showToast`/`guard`/tab code; keep behaviour: `const toast = createToast(getElement("toast"))`, `const guard = makeGuard({ setBusy, notify: toast.show, describe: (e, a) => describeError(e, { host: "Excel", version: formatVersion(__APP_VERSION__) }, a), after: refreshSelection, decorate: (m) => (lastUndoSkipped() ? `${m} (too large for undo)` : m), finally: renderActionState })`; ribbon commands call `describeError` too; `installErrorReporting` at boot; `installTabs(getElement("tab-bar"))` (give the nav an id). Button clicks pass the action id: `guard(() => dispatch(action), action)`.

- [ ] **Step 4: Verify in the browser, not only in tests (lesson 2026-08-27)**

`npm run dev` and open `https://localhost:3000/taskpane.html` in a browser: tabs switch, a forced error (click an action with no Excel) shows the toast with "Copy details". `npm test` green (219 + new), `npm run build` green.

- [ ] **Step 5: Commit**

```bash
git add src/ui src/main.ts taskpane.html src/styles.css
git commit -m "shared pane ui: toast with copy details, guard, tabs, error reporting"
```
