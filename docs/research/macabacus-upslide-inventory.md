# Macabacus + UpSlide feature inventory (research agent, 2026-08-27)

Source: web sweep of help.macabacus.com (via helpscoutdocs mirror) and support.upslide.net.
Shortcuts marked "per recalcacademy.com" are third-party sourced, indicative not authoritative.

# Macabacus

## Excel — Formatting Cycles
| Feature | What it does | Default shortcut | Source URL |
|---|---|---|---|
| Font Color Cycle | Cycles the active cell/selection's font through the user's AutoColor scheme colors with repeated keystrokes | Ctrl+Shift+C (per recalcacademy.com) | https://help.macabacus.com/article/771-format-cycles |
| Fill/Shading Color Cycle | Cycles cell fill/shading color through a defined palette on repeated keystrokes | Ctrl+Shift+V (per recalcacademy.com) | https://help.macabacus.com/article/771-format-cycles |
| Border Color Cycle / Border Style Cycles | Applies border styles using the configured default border color; also used by the Sum Bar tool | Ctrl+Shift+B top border / Ctrl+Shift+7 outer (per recalcacademy.com) | https://help.macabacus.com/article/779-borders |
| Pinstripes (Alternate Row/Column Shading) | Shades odd/even rows or columns in the selection using conditional formatting and the default shading color | - | https://help.macabacus.com/article/771-format-cycles |
| Number Format Cycle | Cycles curated financial-modeling number formats (general, currency, percent, positive/negative) | Ctrl+Shift+N general / Ctrl+Shift+P percent (per recalcacademy.com) | https://help.macabacus.com/article/777-number-formats |
| Left Indent Cycle | Repeating the keystroke increases left indent up to a configured max, then wraps to zero | - | https://help.macabacus.com/article/780-alignment |
| Right Indent Cycle | Same as Left Indent Cycle, applied to right-side indentation | - | https://help.macabacus.com/article/780-alignment |
| Center Alignment Cycle | Cycles centering/alignment options for the selection | Ctrl+Shift+M (per recalcacademy.com) | https://help.macabacus.com/article/780-alignment |
| Underline Cycle | Cycles underline styles (incl. accounting-style underlines used for header rows) | Ctrl+Shift+U (per recalcacademy.com) | https://help.macabacus.com/article/771-format-cycles |
| Header / Date-Header style shortcuts | One-touch "centered header with accounting underline" and "date header" styles | Ctrl+Shift+F header / Ctrl+Shift+H date header (per recalcacademy.com) | https://help.macabacus.com/article/771-format-cycles |
| Custom Styles / Style Cycles | User-defined, reusable multi-property styles (font, color, number format, cell size, text case, validation, comments) grouped into up to 8 "Style Cycles" triggered by one keystroke | Customizable | https://help.macabacus.com/article/783-custom-styles |
| Paintbrush | Copies/applies cell formatting without using the Windows clipboard; stores multiple styles (FIFO) independent of copy/cut operations | - | https://help.macabacus.com/article/784-paintbrush |
| Toggle Formatting | Generic formatting toggle shortcut | Ctrl+Shift+6 (per recalcacademy.com) | https://help.macabacus.com/article/771-format-cycles |

## Excel — AutoColor
| Feature | What it does | Default shortcut | Source URL |
|---|---|---|---|
| AutoColor scheme | Auto-classifies and colors cell fonts by content type: numeric inputs, partial inputs (formulas w/ hardcodes), same-sheet formulas, cross-sheet/workbook formulas, hyperlinks, and external-data-function formulas (FactSet/CapIQ pulls) | N/A | https://help.macabacus.com/article/776-format-colors |
| AutoColor on Entry | Automatically colors fonts per the AutoColor scheme as values/formulas are typed; off by default (perf/Undo impact) | N/A | https://help.macabacus.com/article/776-format-colors |
| Default Font Color | Sets the color used when optionally recoloring fonts on export to PowerPoint/Word/PDF (default black) | N/A | https://help.macabacus.com/article/781-fonts |

## Excel — Footnotes
| Feature | What it does | Default shortcut | Source URL |
|---|---|---|---|
| Footnote Cycle | Formats the selected cell's number format to show a superscripted footnote, incrementing 1-9 on repeated keystrokes | - | https://help.macabacus.com/article/785-footnotes |
| Footnote Toggle | Uses font formatting to superscript a number appended to cell text; repeat to remove | - | https://help.macabacus.com/article/785-footnotes |
| Footnote Hide/Show | Hides footnotes by matching text color to the cell background, or reveals them | - | https://help.macabacus.com/article/785-footnotes |
| Footnote Checker | Validates footnote sequencing/positioning (lower-numbered footnotes must appear in top-left cells) | - | https://help.macabacus.com/article/785-footnotes |

## Excel — Navigation, Selection & Worksheet Tools
| Feature | What it does | Default shortcut | Source URL |
|---|---|---|---|
| Super Find | Replacement for native Find; searches values/formulas/comments/hyperlinks across selection, sheets, workbook, or all open workbooks, with a results task pane and bulk select-and-format | - | https://help.macabacus.com/article/793-super-find |
| NavAid | Shades the rows/columns intersecting the selected range for visibility while navigating | - | https://help.macabacus.com/article/913-excel-performance |
| Sheet Navigation (First/Next/Previous/Last Sheet) | Jumps to first/last sheet, or loops forward/backward through sheets (wrapping at ends) | - | https://help.macabacus.com/article/838-sheets |
| Activate Sheet | Dialog to jump directly to a sheet via search | - | https://help.macabacus.com/article/838-sheets |
| Move Sheets Left/Right | Shifts selected sheet tabs left/right, wrapping at workbook ends | - | https://help.macabacus.com/article/838-sheets |
| Bury/Dig Up Sheets | Hides sheets so they require the add-in to unhide | - | https://help.macabacus.com/article/838-sheets |
| Unhide Sheets (multi) | Enhanced unhide dialog supporting multiple sheets at once with search | - | https://help.macabacus.com/article/838-sheets |
| Resize to Standard Size | Proportionally resizes column widths/row heights of selected cells to a pre-defined Standard Size | - | https://help.macabacus.com/article/791-rows-columns |
| Reverse Columns / Reverse Rows | Reorders columns/rows (e.g. oldest-to-newest periods) while preserving formula references | - | https://help.macabacus.com/article/791-rows-columns |
| Modify Rows / Modify Columns | Batch formatting, insertion, and deletion across multiple rows or columns at once | - | https://help.macabacus.com/article/791-rows-columns |

## Excel — Formula Auditing
| Feature | What it does | Default shortcut | Source URL |
|---|---|---|---|
| Trace In ("Pro Precedents") | Tree-based navigation dialog listing every precedent cell/range of a formula; drill down with arrow keys, edit via F2, evaluate functions/groups | Ctrl+E evaluate within dialog | https://help.macabacus.com/article/842-precedents-dependents |
| Trace Out | Inverse of Trace In — traces which cells depend on the selected cell | - | https://help.macabacus.com/article/842-precedents-dependents |
| Show All Precedents | Shows precedent trace arrows for every selected cell at once (limited to <20 cells) | - | https://help.macabacus.com/article/842-precedents-dependents |
| Formula Flow | Visualizes formula structure via shading/line patterns (horizontal/vertical/crosshatch = consistent; darker = inconsistency) | - | https://help.macabacus.com/article/787-visualizations |
| Dependency Density | Shades cells orange by relative number of dependents, to visualize link density | - | https://help.macabacus.com/article/787-visualizations |
| Uniformulas | Selects/highlights the area of formulaic consistency for the active cell's formula, flagging structural deviations | - | (search snippet; Visualize menu) |
| Model Check (beta) | 50+ automated checks across formula errors, structural issues, hidden data, brand compliance; severity levels, one-click fix/ignore; Professional & Enterprise plans, v9.7.3+ | - | https://help.macabacus.com/article/1216-model-check |
| Formulate | AI feature that replaces hardcoded values with AI-generated formulas; runs inside the firm's own Azure AI tenant | - | https://macabacus.com/features/formulate (snippet) |
| XY Scatter Labels | Auto-detects and applies correct data labels to XY scatter chart points | - | https://help.macabacus.com/article/792-chart-tools |

## Excel — Financial Modeling Tools
| Feature | What it does | Default shortcut | Source URL |
|---|---|---|---|
| Summary Statistics | Auto-computes and inserts min/max/mean/median under a dataset; adapts number formatting, customizable stats | - | https://help.macabacus.com/article/841-financial-modeling |
| Add Scenarios | Builds multiple projection scenarios (upside/mgmt/downside) with local or global (toggle-cell-driven) scope | - | https://help.macabacus.com/article/841-financial-modeling |
| Replicate Module | Duplicates a module/analysis block across worksheets with automatic row insertion; can consolidate copies back | - | https://help.macabacus.com/article/841-financial-modeling |
| Quick CAGR | Inserts a CAGR formula instantly without manually counting periods | - | https://help.macabacus.com/article/841-financial-modeling |

## Excel — Charts
| Feature | What it does | Default shortcut | Source URL |
|---|---|---|---|
| Quick Charts — Waterfall | Builds a waterfall chart from an intuitive data structure; remembers last-used settings/colors; formattable connectors | - | https://help.macabacus.com/article/789-quick-charts |
| Quick Charts — Football Field | Floating-bar valuation-range charts from min/max data | - | https://macabacus.com/blog/build-football-field-chart-excel (snippet) |
| Growth Arrow (Chart Add-on) | Dynamic, data-driven CAGR arrow on stacked/clustered column charts; auto-recalculates | - | https://help.macabacus.com/article/790-chart-add-ons |
| Fast Format (beta) | One-click reformat of any chart to brand-compliant standards; Enterprise only | - | https://help.macabacus.com/article/792-chart-tools |

## Excel — Collaboration & Workbook Tools
| Feature | What it does | Default shortcut | Source URL |
|---|---|---|---|
| Discussions (beta) | Cell/range-anchored chat threads for model review with attachments; survives structural changes | - | https://help.macabacus.com/article/788-discussions-beta |
| Undo/Redo (custom stacks) | Macabacus-specific Undo/Redo stacks that survive add-in operations | Ctrl+Z / Ctrl+Y | https://help.macabacus.com/article/916-undo-redo |
| Soft Disable (Pause/Resume) | One-click pause of most functionality for perf troubleshooting | ribbon | https://help.macabacus.com/article/845-soft-disable |
| Name Scrubber | Shows hidden defined names and safely deletes broken/unused/hidden ones without #REF! errors | - | https://help.macabacus.com/article/797-file-operations |
| Style Scrubber | Bulk-deletes accumulated unused Excel cell styles | - | https://help.macabacus.com/article/797-file-operations |
| Performance / Workbook Optimizations | Diagnostics + one-click optimizations for slow workbooks | N/A | https://help.macabacus.com/article/913-excel-performance |
| Publishing — Send Workbook / PDF to Folder | Converts print areas to PDF and attaches to Outlook email or saves beside workbook; optional font recoloring, error pre-scan | - | https://help.macabacus.com/article/794-publishing |

## Excel-to-PowerPoint/Word Linking & Link Manager
| Feature | What it does | Default shortcut | Source URL |
|---|---|---|---|
| Link mechanism (named ranges) | Links built on a unique hidden range name (not cell address) so links survive row/column changes; chart links key off Selection-Pane name | N/A | https://help.macabacus.com/article/799-link-to-excel |
| Manage Links (Link Manager) | Batch dialog: refresh all, view source, edit/reassign source, break links, Find & Replace file paths | - | https://help.macabacus.com/article/799-link-to-excel |
| Link Autodetection | Resolves link source by open workbooks first, then original file path | N/A | https://help.macabacus.com/article/799-link-to-excel |
| Link Health / Version Control (file-based) | Scans folders for higher-versioned/newer files ("Model_v4.xlsx") and prompts re-linking | N/A | https://help.macabacus.com/article/799-link-to-excel |
| Export as Image / Table / Embedded Workbook / Text / Chart | Five export modes for ranges into PPT/Word with documented tradeoffs | - | https://help.macabacus.com/article/847-import-export-excel-to-powerpoint-word |
| Export charts as Graphic (SVG) / Chart / Picture (EMF) / Embedded Workbook | Chart export modes; SVG recommended | - | https://help.macabacus.com/article/847-import-export-excel-to-powerpoint-word |
| Export settings | Recolor Fonts, Remove Gridlines, Remove Outer Border, Scale Chart Fonts, Check Formula Errors, Copy Appearance, Switch to Target, Placement Options | N/A | https://help.macabacus.com/article/847-import-export-excel-to-powerpoint-word |
| Quick Export — Match Width / Height / Size / None | Four sizing presets for one-click export to a destination shape | configurable | https://help.macabacus.com/article/847-import-export-excel-to-powerpoint-word |
| Paste Exact / Paste Duplicate / Paste Transpose | Excel-side paste ops preserving exact formula text, duplicating with adapted references, transposed paste-linking | - | (third-party confirmation only) |

## PowerPoint — New Presentations & Templates
| Feature | What it does | Source URL |
|---|---|---|
| New Presentation | Creates deck from published template with standard slides + flysheets | https://help.macabacus.com/article/803-new-presentations |
| Rider | Single-slide presentation from template's first content layout | https://help.macabacus.com/article/803-new-presentations |
| Presentation Templates / Template Wizard | 9-step guided .potx configuration (special layouts, agenda placeholders, MasterShapes, validation) | https://help.macabacus.com/article/949-presentation-templates |

## PowerPoint — Agendas & Deck Structure
| Feature | What it does | Source URL |
|---|---|---|
| Agendas (TOC + Flysheets + Section Titles) | Synchronizes TOC slide, section flysheets, on-slide section titles with native sections; auto-updates | https://help.macabacus.com/article/802-agendas |
| Flysheet styles — Topic vs Agenda | Topic = name/number only; Agenda = full TOC with current section highlighted | https://help.macabacus.com/article/802-agendas |
| Subsections | Prefix native section name with "@" | https://help.macabacus.com/article/802-agendas |

## PowerPoint — Shapes, Styles & Brand Compliance
| Feature | What it does | Source URL |
|---|---|---|
| MasterShapes | User-defined native shapes (footnotes, stamps, takeaways, file paths, date/time); insert/update/edit/toggle stamps | https://help.macabacus.com/article/805-mastershapes |
| Styles (New Style / Reset Style) | Reusable shape styles via hidden Slide Master shapes ([STY] suffix) | https://help.macabacus.com/article/807-styles |
| TurboShapes | Dynamic shapes: Harvey Ball, progress bar, rating bar, thermometer, traffic light, toggle, arrow, checkbox, notices + custom | https://help.macabacus.com/article/810-turboshapes |
| Template Rules | Admin compliance rules for required slides and legal-notice text; warn or block on save | https://help.macabacus.com/article/804-template-rules |

## PowerPoint — Deck Check & Version Control
| Feature | What it does | Source URL |
|---|---|---|
| Deck Check — Warnings View | 100+ checks (duplicated phrases, missing parens, orphaned footnotes, bullet punctuation, image distortion, stale legal notices) with auto-fixes | https://help.macabacus.com/article/809-deck-check |
| Deck Check — Reformat View | Inventories font/fill/border/paragraph formatting deck-wide; find-and-replace formatting | https://help.macabacus.com/article/809-deck-check |
| Version Control (library content) | Scans open files for outdated library-sourced content, one-click update | https://help.macabacus.com/article/808-version-control |

## PowerPoint — Slide Tools & Printing
| Feature | What it does | Source URL |
|---|---|---|
| Finalize Pagination — Duplex/Simplex | Inserts blank slides so flysheets/facing slides print correctly | https://help.macabacus.com/article/812-slide-tools |
| Mark as Facing Slide | Flags a slide as "facing" for pagination | https://help.macabacus.com/article/812-slide-tools |
| Slide Numbering (auto-correct) | Start-at-1 or continue; auto-corrects on add/remove | https://help.macabacus.com/article/812-slide-tools |
| Print-Ready Export | Saves a print-ready copy alongside the original | https://help.macabacus.com/article/812-slide-tools |

## PowerPoint — Privacy, Meta Content & Tombstones
| Feature | What it does | Source URL |
|---|---|---|
| Airplane Mode | Redacts logos across all open decks for privacy in public | https://help.macabacus.com/article/815-airplane-mode |
| Meta Content | Publishes slides/shapes/decks to libraries with structured metadata for filter/sort/search; UID-based bulk pairing with Excel data | https://help.macabacus.com/article/817-meta-content |
| Create Templated Shape (Tombstone Generator) | Templated shapes (deal tombstones) via wizard or bulk Excel import; Enterprise-only | https://help.macabacus.com/article/962-tombstone-generator |

## Word
| Feature | What it does | Source URL |
|---|---|---|
| Link to Excel / Export to Word | Same Link Manager/export mechanism as PPT | https://help.macabacus.com/article/799-link-to-excel |
| Libraries (Word) | Document templates, shared text snippets, org content | https://help.macabacus.com/article/816-libraries |
| Doc Builder | Questionnaire-driven document assembly (yes/no keep sections, multiple-choice boilerplate, text placeholders, autofill) | https://help.macabacus.com/article/860-doc-builder |

## Shared Library, Templates & Brand (cross-application)
| Feature | What it does | Source URL |
|---|---|---|
| Libraries (personal & shared) | Charts/tables/text/templates (Excel), slides/shapes/images/decks (PPT/Word); network/cloud folders with permissions | https://help.macabacus.com/article/816-libraries |
| Color Palettes | Unlimited brand palettes across Excel/PPT/Word; "Import Colors" from a cell range; shareable org-wide | https://help.macabacus.com/article/766-color-palettes |
| Standard Sizes | Conforms cells/charts/shapes to preset dimensions | https://help.macabacus.com/article/767-standard-sizes |
| Cloud Sync Folder | Anonymizes user-specific sync paths to %CLOUD_FOLDER% so shared libraries/links resolve across users | https://help.macabacus.com/article/948-cloud-sync-folder |

## AI Features
| Feature | What it does | Source URL |
|---|---|---|
| AI Writing Assistant (AIWA) | Summarize/Elaborate/Rephrase/Proofread/Translate in PPT; enforces Corporate Dictionary/Tone; Azure AI Foundry | https://help.macabacus.com/article/1231-aiwa |
| Corporate Dictionary | Firm-approved terminology via term substitutions and regex rules; used by Deck Check and AIWA | https://help.macabacus.com/article/1230-corporate-dictionary |
| Formulate | AI hardcode-to-formula replacement in firm's Azure tenant | https://macabacus.com/features/formulate (snippet) |

## Settings, Customization & Shortcuts
| Feature | What it does | Source URL |
|---|---|---|
| Shortcut Manager | Lists all shortcuts; Edit, Reset, Clear All, Override (win conflicts), Print | https://help.macabacus.com/article/917-keyboard |
| Accelerator Key | Activates the Macabacus ribbon tab (Alt then B, customizable) | https://help.macabacus.com/article/917-keyboard |
| Disabled Keys | Disable nuisance native keys (F1, Insert, Num Lock, Scroll Lock) | https://help.macabacus.com/article/917-keyboard |
| Manage Settings | Central dialog; export/import XML; reset | https://help.macabacus.com/article/769-manage-settings |
| Share Settings | Publish standardized settings org-wide; auto-download on launch | https://help.macabacus.com/article/770-share-settings |

## Pricing / Editions
- Professional plan: includes Model Check. Enterprise: adds Fast Format, Tombstone Generator; volume discounts >25 licenses.
- Legacy "Modeler"/"Suite" plans exist (excluded from Model Check).

---

# UpSlide (help-center sweep; see also docs/research/upslide-*-training-dump.md primary sources)

## Excel-to-PowerPoint/Word Link & Refresh
| Feature | What it does | Shortcut | Source URL |
|---|---|---|---|
| Create Link / Export (Image, Table, Text) | Exports Excel charts/tables/text into PPT/Word as a live link | - | support.upslide.net/hc/en-us/articles/360025126753 |
| Refresh Selected / Slide / All | Three refresh scopes | - | same |
| Excel Link Manager | Groups links by slide or source; filters by slide/type/last editor/update time; batch updates; flags "Multiple Sources" conflicts | - | support.upslide.net/hc/en-us/articles/11212972252316 |
| View link info / change source / versioning | Inspect source, redirect to different source, file versioning | - | articles/360015936419, /360013569620 |
| Advanced Export — Preserve Cells Visibility | Keeps hidden/grouped visibility state per export; one table linked multiple ways | - | articles/360013548739 |
| Advanced Export — Preserve Image Width | Exported table width expands with added columns instead of shrinking font | - | articles/360013548739 |
| Sizing Guide | Preview target PPT placeholder size before exporting | - | articles/360020079279 (snippet) |
| Data Pack generation | Packaged set of linked exports in one action | - | articles/23280406024860 (snippet) |
| Export Excel table as PPT table | Editable-table export preserving PPT-side formatting on refresh | - | (category listing) |

## Power BI
| Feature | What it does | Source URL |
|---|---|---|
| Power BI to PowerPoint/Word Link | Exports visuals or full pages as live refreshable links | articles/360020765879 |
| Edit Power BI slicers from PPT/Word | Adjust slicer values without returning to Power BI | (snippet) |
| Embed Power BI visuals/pages | Embed mode | articles/360013627180 |
| Translate Power BI dashboards | One-click translation into 44 languages | (snippet) |

## Content & Slide Library
| Feature | What it does | Source URL |
|---|---|---|
| Library (Shape / Slide / Worksheet insert) | Inserts approved content from central library | articles/360015692500 |
| Library navigation & Favorites | Browse/search, favorites | articles/360015709859 |
| Dynamic Library (Bios, CVs, Tombstones, Credentials, Case Studies) | Metadata-filterable content database with auto-arrange into layouts | articles/360019901719 |
| Logo Finder | Search/insert high-res company logos by name incl. multi-search; layout pane; request missing logos | articles/12833661620508 |

## Formatting & Brand Compliance (PowerPoint)
| Feature | What it does | Shortcut | Source URL |
|---|---|---|---|
| Smart Align | Aligns objects to any other shape/placeholder | Ctrl+Alt+Arrows | articles/16934218770716 |
| Align Centre/Middle, Distribute, Resize & Distribute | Centering, even spacing, auto-resize within reference area | - | same |
| Copy Height/Width/Dimensions, Smart Painter | Applies one shape's size or position+size to others | - | same |
| Select Similar Shapes, Arrange, Swap Objects, Swap Text, Group, Rotate, Z-order, Selection Pane | Multi-shape toolkit | - | same |
| Margins, Switch Fonts, Font Size +/-, Footnotes/Endnotes, Paragraph, Text Alignment, Merge Text Boxes | Text-box formatting toolkit | - | same |
| Color Toolbar | Applies brand fill/outline/font color | - | articles/15302629617180 |
| On-brand bullets | Brand bullet styles | - | articles/212067586 |
| Slide Converter | Changes all slide layouts and colors at once | - | articles/360015829759 |
| Facing Pages | Facing-page (spread) layouts | - | articles/360016044880 |
| Branded formatting for tables/charts (Excel) | Brand styling for Excel tables and charts in one action | - | articles/201919038 |

## Excel Modeling & Paste Tools
| Feature | What it does | Shortcut | Source URL |
|---|---|---|---|
| Preserve Formulas | Pastes exact values/formulas keeping original references | Ctrl+Alt+P | articles/360015818779 |
| Duplicate Formulas | In-range refs adapted, absolute/out-of-range preserved | Ctrl+Alt+B | same |
| Row Heights (paste) | Pastes only row heights | Ctrl+Alt+H | same |
| Number Formats Only (paste) | Pastes only number formats | Ctrl+Alt+F | same |
| Fast Fill Down / Right | Auto-fills without pre-selecting destination | Ctrl+Alt+D / Ctrl+Alt+R | same |
| IfError | Wraps formula in IFERROR (or removes) | Ctrl+Shift+I | same |
| CAGR Formula | Auto CAGR on selected data | Ctrl+Shift+Q | same |
| Advanced Charts — Waterfall, Marimekko, Stacked Waterfall | Chart builders auto-linked to the model | - | articles/360015747120 |
| CAGR Arrow (chart) | CAGR growth arrow on an Excel chart | - | articles/4404701072402 |
| PDF to Excel | Converts PDF data into structured Excel | - | articles/28819992564636 |
| Excel Collaboration | Quick collaboration tools | - | articles/360015888600 |
| Audit & present / Clean & prepare workbooks | Modeling-audit and cleanup set | - | articles/360015796100, /360015871780 |

## Agenda / TOC
| Feature | What it does | Source URL |
|---|---|---|
| Insert Table of Contents | Auto-generated TOC slide; auto-splits if long; requires UpSlide template | articles/360015692360 |
| Sections / Subsections | Divider slides with numbering, unnumbered/appendix options | same |
| Auto-update on structural change | Refresh re-syncs numbers, cross-refs, footnotes, breadcrumbs | same |
| Cross-references | References that stay accurate on restructure | articles/360015692380 |
| Footnotes / Endnotes (PPT) | Insert and refresh | articles/10765182231196 |

## Slide Check / Proofing
| Feature | What it does | Source URL |
|---|---|---|
| Slide Check | Flags double spaces, empty placeholders, highlighted text, misaligned shapes, outdated library content, punctuation, non-compliant fonts/styles; per-item or bulk fix | articles/360015692540 |
| AI Consistency Check | AI scan for calculation errors and cross-slide numeric contradictions; flags via comments | articles/24884682152604 |
| Track Changes (PPT) | Diffs two deck versions (SharePoint history or file); accept/reject; marked-up PDF export | articles/19723945482652 |
| Change presentation language | Switches proofing/UI language | articles/360015692560 |
| Finalize presentation | End-of-process finalize | articles/360015709959 |

## Word / Outlook / AI / Admin
| Feature | What it does | Source URL |
|---|---|---|
| Word: Templates, Library, Smart Check, number-to-text, Excel linking | Word toolkit | articles/360016758079 |
| Outlook Signature Manager | Standardized brand signatures firm-wide | articles/9171903202460 |
| AI Assistant / AI Glossary / Claude integration | AI in ribbon; Claude for PPT and Excel via UpSlide | articles/18834156713756, /16735132322076, /28393117120796 |
| Personalized proposal generator | Assembles proposal from library in a few clicks | articles/360010169999 |
| Gantt Charts | Create/edit Gantt charts | articles/23265757975836 |
| Shortcut customization | Ctrl+Shift+[key] primary and Ctrl+Alt+[key] secondary families; reassign/disable | articles/202402408 |
| License portal / deployment / security docs / UpSlide MCP / VBA hook | Admin & IT surface | (IT Corner category) |

---

# Coverage gaps (from the agent, verbatim summary)

- macabacus.com marketing site + PDF shortcut cheat sheets blocked (403); help center read via helpscoutdocs mirror.
- Official default shortcuts largely undocumented in article text (icons, not text); concrete combos are third-party (recalcacademy.com) — indicative only.
- Full Model Check (50+) and Deck Check (100+) rule lists never enumerated anywhere fetched.
- Full Quick Charts catalog beyond Waterfall/Football Field unnamed in docs.
- No Macabacus workbook-diff feature found (Version Control is PPT library-content sync) — likely does not exist.
- Paste Exact/Duplicate/Transpose confirmed only via third parties.
- Macabacus pricing tier mapping pieced from in-article mentions + aggregators.
- UpSlide Outlook/AI/Claude articles seen as titles only, not deep-fetched.
