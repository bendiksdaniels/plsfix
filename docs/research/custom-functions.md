# Custom functions research: `=SMT.ROUND` / `=SMT.ROUNDSUM` (research agent, 2026-08-29)

Backlog item C4 (docs/FEATURES.md section 11). Can SMT ship think-cell TCROUND-style consistent rounding as Excel custom functions inside the shared runtime it already has (SharedRuntime 1.1, `SMT.Taskpane.Url`)? All facts below are learn.microsoft.com (WebFetch, fetched in full, listed at the bottom), except two OfficeDev sample manifests linked directly FROM the `custom-functions-overview` page (ground truth for exact XML wiring) and one npm-registry lookup (webpack-vs-standalone for the JSON generator). Research only.

## Decision table

| Question | Verdict |
|---|---|
| Fits the existing shared runtime? | Yes, additively: no second runtime |
| Do Win/web/Mac (the in-house fleet) clear the requirement sets? | Yes, by years |
| `functions.json` without adopting webpack? | Yes: hand-write it from a typed spec |
| Can Vite emit the separate script Office needs? | Yes: a second, lib-mode config (IIFE) |
| "Sum of rounded parts == rounded sum", achievable? | Yes: stateless recompute per call |
| M365 redeploy needed? | Yes: manifest change, admin re-upload + Update |
| **Recommendation** | **GO** |

## 1. Requirement sets and platform floor

| Requirement set | Web | Windows (M365 sub.) | Windows (retail perpetual) | Mac |
|---|---|---|---|---|
| CustomFunctionsRuntime 1.1 | Supported | 1903 (11425.20156) | 2311 (17029.20126) | 16.34 (20020900) |
| SharedRuntime 1.1 | Supported | 2002 (12527.20092) | 2002 (12527.20092) | 16.35 (20030802) |

Both sets are satisfied wherever SMT already runs; SharedRuntime is the later (binding) build on Windows/Mac, and both floors are ~2020-era. iPad and volume-licensed/LTSC are unsupported, which is irrelevant here: centralized deployment (launch-path.md) only targets M365-subscription tenants, and SMT already excludes iPad.

One doc inconsistency, not actionable: `custom-functions-overview` says in prose that volume-licensed perpetual "2021 or earlier" isn't supported, while both requirement-set tables list "Office 2021: Version 2108" as the supported LTSC floor. Moot for an M365-subscription fleet either way.

## 2. Manifest additions

`<ExtensionPoint xsi:type="CustomFunctions">` needs `Script`/`Page`/`Metadata` (required) and `Namespace` (optional) children, and is Excel-only, so it attaches only to `WORKBOOK_HOST`. It sits in a new `<AllFormFactors>` sibling of `<DesktopFormFactor>`, confirmed against two current OfficeDev sample manifests (`excel-shared-runtime-global-state`, `excel-shared-runtime-scenario`, both linked from `custom-functions-overview`), which order `Runtimes` -> `AllFormFactors` -> `DesktopFormFactor` as direct children of `Host`.

**No second runtime.** Both samples declare exactly one `<Runtime lifetime="long">`, and point the CustomFunctions `<Page>` at that same resid, which for SMT is the existing `SMT.Taskpane.Url`, exactly what the [shared-runtime walkthrough](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime) directs when adding custom functions to an add-in that already has a pane runtime. `<Script>` stays a separate resid pointing at a dedicated `functions.js` file (not the pane HTML) in the cleaner of the two samples. Neither sample declares a `<Set Name="CustomFunctionsRuntime">` inside `<Requirements>`: that capability entry only appears in the *unified JSON* manifest example on the `custom-functions-json` page, not the legacy XML path `manifest/xml.ts` renders; Office gates on the extension point's presence, not a requirement-set declaration.

New fragment `hostBlock()` needs to emit for `WORKBOOK_HOST`, between `<Runtimes>` and `<DesktopFormFactor>` (three new resids: one script URL, one metadata URL, one namespace string; everything else reuses `manifest/spec.ts`):

```xml
<AllFormFactors>
  <ExtensionPoint xsi:type="CustomFunctions">
    <Script>
      <SourceLocation resid="SMT.Functions.Script.Url"/>
    </Script>
    <Page>
      <SourceLocation resid="SMT.Taskpane.Url"/>
    </Page>
    <Metadata>
      <SourceLocation resid="SMT.Functions.Metadata.Url"/>
    </Metadata>
    <Namespace resid="SMT.Functions.Namespace"/>
  </ExtensionPoint>
</AllFormFactors>
```

Plus in `<Resources>`: `bt:Url` ids `SMT.Functions.Script.Url` (`{baseUrl}functions.js`) and `SMT.Functions.Metadata.Url` (`{baseUrl}functions.json`), and `bt:String` id `SMT.Functions.Namespace` = `"SMT"`, matching the ribbon's existing `SMT_AUTOCOLOR`-style prefix, so `=SMT.ROUND(...)` reads consistently with the rest of the add-in.

## 3. functions.json and the script file

Microsoft's own JSDoc tool, `custom-functions-metadata-plugin`, is a **webpack** plugin. Its underlying parser, `custom-functions-metadata` (npm, confirmed directly against the registry: latest 2.2.3), ships its own bundler-agnostic CLI (`bin: custom-functions-metadata generate <src> [out]`), so autogeneration doesn't strictly need webpack. Even so, for exactly two functions, hand-writing `functions.json` from a small typed spec (`manifest/functions-spec.ts` -> `scripts/build-functions-json.ts`, `--check` wired into `npm run check`) fits better: it mirrors the repo's own `manifest/spec.ts` -> `scripts/build-manifests.ts` pattern instead of adding a second, differently-shaped generator and a webpack-flavored dependency to a Vite project.

```json
{
  "functions": [
    { "id": "ROUND", "name": "ROUND",
      "description": "Per-cell rounding: the range's rounded values sum to its rounded total.",
      "result": { "type": "number", "dimensionality": "scalar" },
      "parameters": [
        { "name": "range", "type": "number", "dimensionality": "matrix" },
        { "name": "index", "type": "number", "dimensionality": "scalar" },
        { "name": "decimals", "type": "number", "dimensionality": "scalar" }
      ] },
    { "id": "ROUNDSUM", "name": "ROUNDSUM",
      "description": "The range's grand total, rounded so it equals the sum of SMT.ROUND over it.",
      "result": { "type": "number", "dimensionality": "scalar" },
      "parameters": [
        { "name": "range", "type": "number", "dimensionality": "matrix" },
        { "name": "decimals", "type": "number", "dimensionality": "scalar" }
      ] }
  ]
}
```

`id` is unqualified; the manifest `<Namespace>` prepends `SMT.`. The script must call `CustomFunctions.associate("ROUND", smtRound)` / `associate("ROUNDSUM", smtRoundSum)` once each; an unassociated function registers but throws `#N/A` (`custom-functions-troubleshooting`).

**Vite:** yes, via a second config. `taskpane.html`/`pptpane.html` are HTML entries (Vite resolves the module graph from their `<script type="module">` tags); `<Script>` instead wants a plain classic script that Excel injects into the shared Page: no ES `import`/`export`. One Vite pass can't emit that alongside the main ESM chunks, so this needs a **second, library-mode config** (`build.lib`, `formats: ["iife"]`), output into the same `dist/` with `emptyOutDir: false` so it doesn't erase the main build.

## 4. The rounding algorithm

**Not possible:** a formula cell only ever writes its own cell. There is no mechanism for one cell's formula to push a result into a sibling cell, so cells cannot "negotiate" a shared allocation live. `@streaming` only pushes repeated results into *its own* cell over time; `@volatile` only forces more frequent recompute. Neither grants cross-cell writes, so TCROUND's literal internal mechanism (shared mutable group state, pushed to members) cannot be replicated as such with stock custom functions.

**What works, and is simplest:** pass the whole range into every sibling cell's formula as an explicit argument. Excel already recalculates a cell whenever any argument in its own formula changes (no `@volatile` needed), so every `SMT.ROUND` cell in the group naturally recalculates whenever *any* cell in the range changes, because each one lists the whole range as a dependency. Each cell then independently recomputes the identical deterministic allocation from identical inputs and reads out its own slot. The group behaves as if it shared live state, without ever needing real shared state or messaging: a stateless, full recompute on every call, exactly as the task's proposed signatures assume.

```
allocate(values, decimals):
  unit      = 10^-decimals
  target    = round(sum(values), decimals)              # the correct grand total
  floors    = values.map(v => floor(v, decimals))
  shortfall = round((target - sum(floors)) / unit)       # whole units still owed
  order     = indices sorted by (remainder desc, index asc)  # largest remainder first, stable
  bumped    = first `shortfall` indices in order
  return values.map((_, i) => floors[i] + (bumped.has(i) ? unit : 0))
```

`SMT.ROUND(range, index, decimals)` returns `allocate(range, decimals)[index-1]`. `SMT.ROUNDSUM(range, decimals)` returns `round(sum(range), decimals)`: definitionally what `allocate` targets; it exists mainly so the total cell is visibly namespace-paired with the per-cell formulas for audit purposes, not because it needs its own logic.

`index` is a plain 1-based integer, not a cell address, supplied as a **literal**, written by an SMT ribbon button when it inserts the formulas (`=SMT.ROUND($A$1:$A$3,1,0)` in A1, `...,2,...` in A2, etc.), the same way `src/paste.ts`'s `buildCagrFormula` is composed by `src/excel/formulas.ts` today. This sidesteps any need for `@requiresAddress` introspection, at the cost that a formula copied to a new row by hand (not via the button) keeps its old literal index: a real but narrow gap, flaggable later by the existing Audit overlay pattern (`src/excel/audit.ts`).

**Worked example.** Allocate 100.0 across 33.333 / 33.333 / 33.334 at 0 decimals:

| cell | value | floor | remainder | rank | bumped | result |
|---|---|---|---|---|---|---|
| A1 | 33.333 | 33 | .333 | 2 | no | 33 |
| A2 | 33.333 | 33 | .333 | 3 | no | 33 |
| A3 | 33.334 | 33 | .334 | 1 | **yes** (shortfall=1) | 34 |
| Total | 100.000 | n/a | n/a | n/a | n/a | **100** = ROUNDSUM |

33 + 33 + 34 = 100, matching `ROUNDSUM`. A3 wins here on a strictly larger remainder; true ties (equal remainders, e.g. splitting 10.00 three ways) fall back to the stable index order above: the reason `index` must carry real positional identity, not be inferred from value.

**Why the shared runtime at all, if the range arrives as a plain parameter?** Excel marshals a range-reference argument into a `number[][]` automatically, so neither function needs `Excel.run()`/`context.sync()` to read the sheet, and strictly speaking neither *needs* the shared runtime. Ship them there anyway: it's Microsoft's default guidance once any shared runtime exists ("always use a shared runtime unless you have a specific reason not to"), reuse is free (same `Page` resid, zero new infrastructure), and it leaves room for a future variant that reads context SMT.ROUND doesn't get as an argument (e.g. "the row above me") without a second migration.

## 5. M365 deployment implications

Adding the extension point changes `manifest.xml`/`manifest.prod.xml` structurally (new `<ExtensionPoint>`, new resources), not just hosted assets. Per docs/research/launch-path.md: regenerate via `npm run manifest:build`, then the M365 admin re-uploads/updates the manifest in Integrated Apps, effective at each user's *next* Excel launch (~1-3 days to fully propagate). Separately, and specific to custom functions, Office caches function registration independently of the manifest: per `custom-functions-troubleshooting`, "changes to the **functions.json** and **functions.js** files may take up to **24 hours** to reach your end users, while changes to taskpane.html reach end users more quickly." An add-in can force an immediate refresh via `Office.context.document.settings.set('Office.ForceRefreshCustomFunctionsCache', true)` (shared-runtime only), but Microsoft warns frequent use "can impact performance"; use it only during the initial rollout window, then disable it.

AppSource is not on the in-house path (launch-path.md already recommends centralized deployment only), so AppSource-specific caveats don't gate this. One caveat applies regardless of distribution channel: coauthoring. A colleague opening a workbook that already contains `=SMT.ROUND(...)` formulas is prompted to load the SMT add-in to see live results (`custom-functions-overview`, Coauthoring section). Expected, not a blocker.

## 6. Risks

**Caching:** up to 24h lag on functions.json/js-only edits after launch (see above); mitigate with the force-refresh flag during rollout only, then disable it. **Volatility:** do not mark either function `@volatile`. Passing the whole range as an argument already keeps the group in sync via normal dependency tracking; `@volatile` would instead force recompute on *every* recalculation regardless of whether the range changed, which is pure waste, and the docs warn it can "make recalculation times slow." **Performance on large models:** each call is O(N log N) (one sort); N sibling `SMT.ROUND` cells in one group means O(N² log N) per full-group recalc, since every cell redoes the whole allocation. Trivial at reconciliation-row scale (a handful to a few dozen cells, the actual TCROUND use case); a real cost in the hundreds+. Worth a documented practical ceiling on range size, and a module-scope memoization cache (keyed by a fingerprint of the range values + decimals) as a follow-up if a large use case appears, available in either runtime since both keep one persistent JS module scope across calls within a session. **Mac support:** not a real risk. Combined floor is Excel Mac 16.35 (SharedRuntime 1.1, the binding constraint since it's newer than CustomFunctionsRuntime 1.1's 16.34), a 2020-era build; any Mac in the in-house fleet clears it by years.

## Build-integration plan

- `manifest/spec.ts`: a `CustomFunctionsSpec` (namespace, script/metadata resource ids) attached only to `WORKBOOK_HOST`. `manifest/xml.ts`: `hostBlock()` emits `<AllFormFactors>` + the extension point when `host.customFunctions` is set; `resources()` gains the two new `bt:Url`s and the namespace `bt:String`.
- `src/rounding.ts` (new, pure, no I/O): `allocate()` above, unit-tested directly: the worked example, the tie-break case, negative values, `decimals > 0`, `N=1`.
- `src/functions/functions.ts` (new): thin `CustomFunctions.associate` wrappers around `src/rounding.ts`; becomes the IIFE entry. Tested by calling the associated functions directly as plain JS, the same way `src/paste.ts`'s pure helpers are tested today; no `test/fakehost.ts` involvement, since neither function touches the Excel object model.
- `manifest/functions-spec.ts` + `scripts/build-functions-json.ts` (new, mirrors `scripts/build-manifests.ts`), `--check` wired into `npm run check`.
- `vite.functions.config.ts` (new): `build.lib` entry `src/functions/functions.ts`, `iife` format, `outDir: "dist"`, `emptyOutDir: false`. `npm run build` becomes `tsc --noEmit && vite build && vite build --config vite.functions.config.ts`.
- Ribbon: one `ButtonSpec` in `SMT.Group.Model` ("Consistent rounding"); `src/excel/formulas.ts` gets `insertConsistentRounding()` writing the literal per-cell `SMT.ROUND` formulas plus one `SMT.ROUNDSUM` into a chosen total cell, fake-host tested like the rest of that file.

## Effort estimate

~6 agent-tasks at this repo's usual granularity (tasks/v2-backlog.md B/C items): (1) `src/rounding.ts` + tests, (2) `functions-spec.ts` + JSON generator + `--check` gate, (3) `manifest/spec.ts`/`xml.ts` CustomFunctions block + regenerate + `npm run validate`, (4) `vite.functions.config.ts` + build wiring + confirm `dist/functions.js` is a plain IIFE script, (5) `src/functions/functions.ts` + direct-call tests, (6) ribbon button + `src/excel/formulas.ts` insertion + fake-host test. Plus one Daniel-gated item outside that count: a real-Office sideload pass (Windows/web/Mac) to confirm registration, IntelliSense text, and #NAME?/cache behavior, the same shape as the "real-Office pass" already sitting in "Blocked on Daniel" in tasks/v2-backlog.md.

## GO / NO-GO

**GO.** Deciding fact: two currently-maintained OfficeDev sample manifests, linked directly from Microsoft's own `custom-functions-overview` page, prove the CustomFunctions extension point reuses the *existing* shared-runtime `Page` resid: no second runtime, no parallel infrastructure. Everything else (functions.json, the Vite script, the algorithm) is ordinary same-repo-pattern work, not a new architectural risk.

## Sources (learn.microsoft.com, fetched in full)

custom-functions-overview, custom-functions-runtime, develop/configure-your-add-in-to-use-a-shared-runtime, javascript/api/requirement-sets/excel/custom-functions-requirement-sets, javascript/api/requirement-sets/common/shared-runtime-requirement-sets, javascript/api/manifest/extensionpoint, custom-functions-json-autogeneration, custom-functions-json, custom-functions-troubleshooting.

Ground truth (linked from custom-functions-overview): OfficeDev/Office-Add-in-samples `excel-shared-runtime-global-state` and `excel-shared-runtime-scenario` manifest.xml. Package facts: registry.npmjs.org for `custom-functions-metadata` / `custom-functions-metadata-plugin`.
