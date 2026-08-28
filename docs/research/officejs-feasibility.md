# Office.js feasibility for a Macabacus-class add-in (research agent, 2026-08-27)

Source: learn.microsoft.com sweep. Rows marked "snippet-sourced" carry lower confidence; see Coverage gaps.

| Capability | Support | API / requirement set | Platforms | Notes & workaround |
|---|---|---|---|---|
| Keyboard shortcuts, core | Yes | SharedRuntime 1.1 + KeyboardShortcuts 1.1; Office.actions.associate / replaceShortcuts / getShortcuts / areShortcutsInUse | Excel, PowerPoint, Word only. Web; Win Excel 2102+; Mac Excel 16.55+ | Requires shared runtime. Shortcuts run manifest-defined executeFunction/ShowTaskpane actions, one modifier + one key. On desktop they run with no pane visible; on web focus must be on the document. Users can remap via replaceShortcuts (signed-in only); prefs per-user per-platform. |
| Keyboard shortcuts, conflicts | Partial | same | Web most restricted | Conflict shows a one-time picker dialog. On web, Ctrl+X/C/V/N, Ctrl+W, Ctrl+PgUp/PgDn and several others can NEVER be overridden. |
| Ribbon add-in commands (tabs/groups/buttons/menus) | Yes | Manifest AddinCommands; icons 16/32/80px | Web, Win + Mac with M365, perpetual 2021+ | One custom core tab per add-in; menus but no submenus; recommended limits 7 tabs / 6 groups/tab / 7 commands/group / 70 total. Content add-ins get no commands. |
| Ribbon function commands (run code, no pane) | Yes | Browser runtime (5-min timeout) or shared runtime (no timeout) | as above | The Office.js equivalent of a Macabacus button that "just does the thing"; can use displayDialogAsync for progress/input UI. |
| Custom contextual tabs | Yes, Excel only | RibbonApi 1.2 + SharedRuntime 1.1; Office.ribbon.requestCreateControls / requestUpdate | Excel web; Win 2102+; Mac 16.53+. Not PPT/Word | Runtime JSON definition; register once per session; needs documented fallback for unsupporting hosts. |
| Enable/disable ribbon controls at runtime | Partial | Office.ribbon.requestUpdate (RibbonApi) | not all hosts/scenarios | Works for core + contextual in one call; Microsoft flags host/version gaps. |
| Shared runtime | Yes | SharedRuntime 1.1, lifetime "long" | Excel/PPT/Word web + modern desktop | THE load-bearing feature: shared in-memory state between ribbon and pane; unlocks contextual tabs, shortcuts, showAsTaskpane/hide, full-CORS custom functions. Only ONE task pane may share the runtime; Dialog never shares it. |
| Range formatting at scale, batching | Yes with discipline | Application.suspendApiCalculationUntilNextSync / suspendScreenUpdatingUntilNextSync; context.runtime.enableEvents=false | all | Batch loads, one sync; write via one 2D array per block; untrack() proxies in 10k+ cell loops; don't loop suspendScreenUpdating (flicker). |
| Range formatting hard limits | bounded | platform limits | web 5MB; 5M-cell read all platforms | Excel web caps payload at 5MB (RichAPI.Error); >5,000,000-cell reads error or return null — chunk reads. |
| Number formats | Yes | Range.numberFormat / numberFormatLocal; Style.numberFormat at ExcelApi 1.7 | all | standard surface. |
| Cell styles / named styles | Yes | Excel.Style + workbook.styles (ExcelApi 1.7; autoIndent/textOrientation 1.8) | all | Full font/fill/borders/number/alignment/protection; StyleCollection.add() not directly confirmed (gap). |
| Workbook themes/fonts API | Not confirmed | - | - | No page found either way; check Excel.Workbook reference directly. |
| Formula auditing precedents/dependents | Yes | Range.getPrecedents / getDirectPrecedents / getDependents / getDirectDependents -> WorkbookRangeAreas | all | Cannot cross workbook boundaries. Direct methods first (full traversal slow). No native-style arrows — draw your own. Throws ItemNotFound. No formula AST API beyond range.formulas. Min ExcelApi version unconfirmed (gap). |
| Paste-special equivalents in-workbook | Yes | Range.copyFrom(source, copyType[all/formulas/values/formats], skipBlanks, transpose); Range.moveTo | all | Faithful Paste Special equivalent range-to-range. |
| OS clipboard interception | No | - | - | No API reads/observes/writes the system clipboard or hooks native paste. Task pane webview navigator.clipboard is sandbox-only. HARD BLOCKER for Macabacus-style "paste exact" over Ctrl+V. |
| Selection/document events | Yes broad | onSelectionChanged, onChanged, onCalculated, onActivated/Deactivated, onFormatChanged, onFormulaChanged, onRowHiddenChanged, onRowSorted/onColumnSorted, onNameChanged, onMoved, onAdded/onDeleted, preview onFiltered | all | onChanged carries address/details, no per-cell bulk-paste granularity. Coauthoring events carry source Local/Remote. Handlers don't persist across reload. Known bug: onRowHiddenChanged misses Advanced Filter hides (poll workaround). Disable events during batch edits (perf lever). |
| Charts incl. waterfall | Yes | Excel.ChartType.waterfall at ExcelApi 1.9; worksheet.charts.add | all with 1.9+ | Waterfall natively creatable (green light for bridge builder). 1.9 also added Boxwhisker, Funnel, Pareto, RegionMap, Sunburst, Treemap. |
| Chart/range image export | Yes | Range.getImage() (ExcelApi 1.9) / Chart.getImage(w,h,fit) (ExcelApi 1.2) -> base64 PNG; Shape.getImageAsBase64 is PowerPointApi 1.10, Slide.getImageAsBase64 1.8 | all (Mac had an upside-down-image bug report; verify in the v2 spike) | Verified 2026-08-28 against learn.microsoft.com. Range.getImage needs no manifest change (1.9 was already the floor). |
| Excel add-in <-> PowerPoint add-in direct comms | No documented channel | - | - | Shared runtime is one add-in in one host; Dialog messaging is own-dialog only. Verdict inferred from absence (gap noted). |
| Excel->PPT image + insertion path | Yes | Excel getImage -> PowerPoint: addGeometricShape (1.4) + shape.fill.setImage(base64) (PowerPointApi 1.8, Win 2504 / Mac 16.96, NOT on VL/LTSC); below 1.8 the Common API setSelectedDataAsync(CoercionType.Image) inserts a picture at the selection. ShapeCollection.addPicture is PREVIEW ONLY (not in shipped typings). Native tables: addTable 1.8, cell/row/column formatting 1.9. | per host | Corrected 2026-08-28: no production image-insert API exists; refresh in place = fill.setImage on a rectangle, geometry untouched. |
| Durable identity on a PowerPoint shape | Yes | PowerPoint.Shape.tags / TagCollection (add = upsert, keys stored UPPERCASE, string values; getItemOrNullObject, delete) on Shape, Slide and Presentation; PowerPointApi 1.3 | all | THE tracker primitive for Excel->PPT links (v2): identity lives in the tags, never in shape id/name/position. Persisted in the .pptx; survival across cut/paste and copy to another deck is undocumented -> verified by the v2 spike (tasks/AUTORESUME.md). Added 2026-08-28. |
| Metadata persistence per host | Yes per-host | Excel.CustomXmlPart (ExcelApi 1.5); PowerPoint.CustomXmlPart (PowerPointApi 1.7); Word via Common API; all hosts: Office.context.document.settings (saveAsync) | all | File-embedded, survives reopen/coauthoring. No cross-file store. Settings size limit unofficial ~1MB reports (gap). |
| "Live link" Excel->PPT needs | Backend or Graph | - | - | No add-in-to-add-in channel + no cross-file metadata => link refresh needs (a) backend service both add-ins call, (b) shared cloud file via Microsoft Graph, or weakly (c) OfficeRuntime.storage shared same-domain same-device (not a real sync). |
| Sheets, hidden sheets, navigation | Yes | Worksheet.visibility (SheetVisibility Visible/Hidden/VeryHidden), read 1.1 write 1.2; onMoved/onNameChanged/onVisibilityChanged | all | Deleting VeryHidden throws until visibility changed. Snippet-sourced. |
| Defined names | Yes | NamedItemCollection.add(name, reference, comment) at ExcelApi 1.4; scope worksheet/workbook; addFormulaLocal | all | Snippet-sourced. |
| TOC generation | No dedicated API | - | - | Hand-build from worksheets + hyperlinks + defined names. |
| Custom functions (UDFs) | Yes | CustomFunctions set; @customfunction JSDoc | Excel web/Win/Mac; NOT iPad or volume-licensed perpetual <=2021 | Namespaced like native (=NS.FN); coauthors prompted to load add-in. |
| Custom functions streaming/runtime | Yes | StreamingInvocation + setResult; shared runtime (full CORS, can read sheet) vs JS-only runtime (fast, simple CORS, no localStorage -> OfficeRuntime.storage 10MB/domain) | as above | |
| Dialog API | Yes with constraints | displayDialogAsync; DialogApi 1.2 messageChild; Dialog Origin 1.1 | all | Nonmodal, HTTPS same-domain default, one dialog per host window, default 80% screen, displayInIframe on web breaks sign-in pages, never window.open(). |
| Task pane UX constraints | Partial | - | all | One shared-runtime pane (multiple HTML views inside); extra ShowTaskpane action = separate non-shared runtime. No documented px width bounds (gap). |
| OfficeRuntime.storage | Yes | 10MB/domain, unencrypted key/value | hosts with custom functions/shared runtime | Shareable between two add-ins on the SAME domain (subdomains separate); device-scoped, not in file; no clear(), use removeItems. |
| document.settings | Yes | Common API Settings, saveAsync into the file | Excel, PPT, Word | JSON k/v in the file, scoped to the add-in. Size limit unofficial (~1MB reports). |
| localStorage | Partial | Web Storage | pane/dialog/full runtimes only | Not in JS-only custom-function runtime; device/profile-scoped, wiped by cache resets. |
| Centralized Deployment | Yes | Integrated Apps portal (recommended) | Win + Mac + web | Needs M365 business/enterprise SKUs, Exchange Online OAuth, admin role; no on-prem Exchange, no SharePoint catalog, no COM/VSTO. Users/groups/org assignment; up to 24h propagation. |
| AppSource vs sideload | Yes | AppSource for public; sideload dev/test only | all | Event-based add-ins need admin deployment or restricted listing to activate. |
| Admin-managed settings | Yes | M365 admin center | all | Enable/disable, assignment, org-wide removal; manifest changes to event-based add-ins need re-consent. |
| Right-click context menus | Partial (narrow) | ExtensionPoint ContextMenuText + ContextMenuCell, OfficeMenu | Word, Excel, PPT, OneNote | Can ADD items to text-selection and Excel cell menus only; cannot add menus to arbitrary objects nor suppress/replace native entries. Snippet-sourced. |
| Unified manifest | Yes, split support | manifest.json vs legacy XML | Web yes; Outlook Win; Excel/PPT/Word Win 2501+, Mac 16.103+; NOT perpetual Win, Outlook Mac, mobile | Carries keyboardShortcuts/runtimes/ribbons/autoRunEvents as first-class JSON. Ship two manifests side-by-side until coverage completes (Microsoft's own guidance). |
| Event-based activation (Excel/PPT/Word) | Yes but thin | autoRunEvents; currently just OnDocumentOpened | Win 16.0.18324.20032+ and web; Mac "later" | JS-only runtime, no UI APIs, ~300s timeout, max 5 active event add-ins, one non-deterministic winner per event if several subscribe; needs admin deployment or restricted listing. |

# Hard blockers (impossible in web add-ins)

- Arbitrary custom right-click context menus (only ContextMenuText/ContextMenuCell insertion points; no suppression/replacement of native entries).
- Intercepting native commands (Ctrl+V, Ctrl+S, Ctrl+P, Undo). Shortcut registration is additive with a user-facing conflict dialog; a fixed list is unoverridable on web.
- True OS clipboard interception. copyFrom is object-model range-to-range only.
- Modifying native Office UI/chrome outside sanctioned extension points; one custom core tab per add-in (contextual tabs Excel-only).
- Reading or acting on OTHER open workbooks/documents. Office.js is scoped to the one host document; auditing APIs refuse workbook boundaries.
- General file-system access (only current document via getFileAsync 4MB slices + sandboxed storage surfaces).
- Application-level object model/events (NewWorkbook, WindowActivate, app-wide before-save/print).
- Direct add-in-to-add-in / cross-host messaging (linking needs a backend or Graph/OneDrive relay).

# Coverage gaps (verify before implementation)

- Workbook theme/font-scheme API existence (check Excel.Workbook reference).
- Min ExcelApi versions for getPrecedents family and Range/Chart.getImage.
- document.settings official size limit (unofficial ~1MB reports).
- Task pane px width bounds.
- StyleCollection.add() (creating new named styles) not directly confirmed.
- PowerPoint Shape/ShapeCollection + Excel getImage mechanics rest on snippets — deep-fetch before building the linking milestone.
- No single Microsoft page states "Excel and PowerPoint add-ins cannot communicate" — verdict inferred from absence across ~40 fetches.

# Key pages consumed (all fetched in full)

learn.microsoft.com: keyboard-shortcuts, contextual-tabs, add-in-commands, precedents-dependents, event-based-activation, shared-runtime configuration, centralized-deployment requirements, custom-functions-runtime, excel performance, Excel.Style reference, ranges cut-copy-paste, excel events, custom-functions-overview, dialog-api, unified-manifest-overview, Excel.ChartType enum, task-pane/content add-in support, resource-limits. Snippet-only: Excel.Range, PowerPoint.SlideGetImageOptions, PowerPoint shapes, Excel/PowerPoint CustomXmlPart, NamedItemCollection, SheetVisibility, ExtensionPoint/OfficeMenu manifest refs.
