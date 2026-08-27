# UpSlide help-center full inventory (research agent, 2026-08-27)

Source: support.upslide.net English help center (~60 articles fetched in full; admin rows marked
"listing only" are inferred from article titles). Note: shortcuts here sometimes differ from the
sales-team training files in docs/research/upslide-*-training-dump.md (e.g. Autocolor vs Smart Track
keys are swapped between the two sources) — UpSlide shortcuts are user-remappable, so treat all keys
as indicative defaults.

## Excel — formatting / Smart Format / custom styles
| Feature | What it does | Shortcut |
|---|---|---|
| Smart Format (tables) | Title style first row, result style last row, item style between, from company theme | Ctrl+Shift+S |
| Toggle custom title formats | Cycles approved title-row formats | Ctrl+Shift+T |
| Toggle custom result formats | Cycles approved result-row formats | Ctrl+Shift+R |
| Toggle custom item formats | Cycles approved item-row formats | Ctrl+Shift+N |
| Toggle custom backgrounds | Cycles brand fill colors | Ctrl+Shift+C |
| Toggle custom borders (Line Styles) | Cycles approved border/line styles | Ctrl+Shift+O |
| Toggle custom fonts | Cycles company typefaces | Ctrl+Alt+T |
| Font Color Cycle | Cycles brand font colors | Ctrl+Alt+Shift+A |
| Font Size set / increase / decrease | Admin-defined sizes | Ctrl+Alt+Shift+Q / F / G |
| Toggle custom number formats | Cycles custom formats by category (General/Currency/Percent/Multiple/Date) with previews | Ctrl+Shift+F |
| Toggle custom percentage formats | Cycles percentage formats | Ctrl+Shift+% |
| Toggle custom row heights / column widths | Cycles preset standards | Ctrl+Shift+H / Ctrl+Shift+W |
| General / Date / Currency / Multiple format cycles | Dedicated per-family cycles | Ctrl+Shift+1 / 2 / 4 / 8 |
| Clear all formats | Strips formatting | Ctrl+Alt+C |
| Change sign of numbers | Flips sign of selected cells | Ctrl+Shift+- |
| Divide / Multiply by 1000 | Rescales values | Ctrl+Shift+/ and Ctrl+Shift+* |
| More / fewer decimal places | Steps decimals | Ctrl+Shift+, / Ctrl+Shift+. |
| Smart Format (charts) + Auto Smart Format | One-click branded chart formatting; auto-applies to new charts | - |

## Excel — charts
| Feature | What it does | Shortcut |
|---|---|---|
| Waterfall / Bridge chart | Cumulative +/- bridge | Ctrl+Shift+B |
| Stacked Waterfall | Waterfall by subcategory, stack/unstack pillars | - |
| Marimekko | 100% stacked with width = second variable | - |
| S-curve | Marimekko variant, raw values | - |
| CAGR Arrow | CAGR between first/last chart points as overlay arrow | Ctrl+Shift+G |
| Update waterfall Y-axis + CAGR | Refresh scales/arrows after data change | Ctrl+Shift+U |
| Select Source Data | Reassign chart range after structural edits | - |
| Chart Examples | Preview types, insert sample data | - |

## Excel — modeling, audit & navigation
| Feature | What it does | Shortcut |
|---|---|---|
| Preserve Formulas paste | Exact values/formulas, original references | Ctrl+Alt+P |
| Duplicate Formulas paste | In-range refs adapt, external absolutes preserved | Ctrl+Alt+B |
| Paste Row Heights | Row-height formatting only | Ctrl+Alt+H |
| Paste Number Formats Only | Number formats only | Ctrl+Alt+F |
| Fast Fill Right / Down | Pattern-aware fill without pre-selecting destination | Ctrl+Alt+R / Ctrl+Alt+D |
| Apply IFERROR | Wraps in IFERROR with custom error value | Ctrl+Shift+I |
| Add CAGR formula | CAGR over range + period (annual/quarterly/monthly) | Ctrl+Shift+Q |
| Insert TOC | Hyperlinked list of visible sheets, auto-updates on rename | - |
| Explorer | Tree pane of open workbooks/sheets with search | Ctrl+Alt+X |
| Autocolor | Color-codes cells by type; Autocolor on Edit lives | Ctrl+Shift+A (help center; training file says Ctrl+Shift+K) |
| Smart Track | Keyboard precedent/dependent navigation with color trail | trace prec Ctrl+Shift+K, dep Ctrl+Shift+L, pane Ctrl+Shift+Space, back Ctrl+Shift+J (help center; training file says pane = Ctrl+Shift+A) |
| Formula Audit | Striped fill = consistent formulas, solid fill = deviation | - |

## Excel — Clean & Prepare
Clean (metadata/size, up to ~95% reduction claimed): Names and Metadata (broken refs), Styles (unused, style-ceiling), View and Range (crop to used range). Prepare (external sharing): formulas -> values, strip comments, delete hidden content, reset zoom.

## Excel — linking & export (to PPT/Word)
| Feature | What it does | Shortcut |
|---|---|---|
| Export to PowerPoint / Word | Linked, one-click-updatable image export | Ctrl+Shift+E / Ctrl+Shift+D |
| Resize and Export | Pre-select PPT placeholder, size to fit | - |
| Export Text | Cell value -> live text in PPT placeholder | - |
| Paste Link | Copy in Excel, paste in PPT/Word -> prompt to make dynamic link | - |
| Hide/Unhide Text Links | Toggle hyperlink styling on text exports | - |
| Excel table as native PPT table | Editable PPT table instead of image | - |
| Advanced Export | Preserve cell visibility; preserve/override image width | Ctrl+Alt+A |
| Highlight linked items | Shows which Excel cells are linked out | Ctrl+Alt+L |
| Sizing Guide | Target placeholder dimensions shown in Excel | - |
| Excel Link Manager | All links grouped by slide/workbook; batch update; filters; bulk-fix broken | Ctrl+Shift+R |
| Change sources & versioning | Repoint links to new/moved file versions | - |
| Link info / Go to Source / last export | Jump between link ends; rename/move safe | - |
| Update / Update All | Refresh one or every link | - |
| Data Pack | Standalone Excel collecting every linked table/chart of a deck, with contents page | - |

## Excel — collaboration & AI
Display Excel library Ctrl+Shift+Y; Send selection/workbook via email (xls/pdf/image/HTML, compression) Ctrl+Shift+M; Smart Print (custom headers/orientation/print areas) Ctrl+Shift+P; PDF to Excel (AI region extraction); AI Assistant Ctrl+Shift+Space.

## PowerPoint — templates, library, structure
| Feature | What it does | Shortcut |
|---|---|---|
| Templates | New deck from approved template | Ctrl+Shift+D |
| Cycle slide layout | Steps through layout variants | Ctrl+Shift+K |
| Proposal Wizard | Auto-assembles customized pitch from content sources | - |
| Library open / insert / favorites | Approved slides/shapes/worksheets/photos | Ctrl+Shift+L |
| Dynamic Library | Filterable bios/CVs/tombstones/credentials/case studies | - |
| Logo Finder | Search/insert logos, bulk + arrangement | - |
| Linked slides | Master slide updates propagate to reuses | - |
| Locked slides | Insertable but not editable (disclaimers) | - |
| Insert TOC | Auto contents page, splits when long | Ctrl+Shift+T |
| Insert Section / Subsection divider | Formatted dividers, TOC auto-update | Ctrl+Shift+S / Ctrl+Shift+B |
| Unnumbered / Appendix sections | TOC numbering exclusions | - |
| View Outline | Structure pane with drag-and-drop reorder | Ctrl+Shift+P |
| Refresh (TOC/dividers/numbers/cross-refs/footnotes) | One-click resync | Ctrl+Shift+U |
| Reminders / Breadcrumbs | Section context on every slide | - |
| Cross-references | Dynamic slide refs (number/section/title) | - |
| Footnote / Endnote | Superscript markers; endnotes collect on auto slide | Ctrl+Alt+F / Ctrl+Alt+G |

## PowerPoint — formatting shapes
Arrange Ctrl+Shift+A; Smart Painter Ctrl+Shift+C; Apply Height/Width Ctrl+Shift+H/W; Copy Dimensions Ctrl+Shift+Z; Smart Align Ctrl+Alt+arrows; Resize & Distribute Ctrl+Alt+D/H; Select Similar Ctrl+Shift+Y; Swap top-left Ctrl+Alt+S; Swap center Ctrl+Alt+W; Swap text Ctrl+Alt+T; Margins default/remove/custom Ctrl+Alt+O / Ctrl+Shift+O / Ctrl+Alt+P; Merge Text Boxes; Color Toolbar Ctrl+Alt+Y; Custom Shapes (+ user-added); Table Styles; Paragraph Styles; Apply Master Format Ctrl+Shift+arrows; Switch body/heading font Ctrl+Alt+B; on-brand bullets.

## PowerPoint — check, review, finalize
Slide Check Ctrl+Alt+C (typos, formatting, alignment, outdated library content, fonts, placeholders); AI Consistency Check (calc errors, cross-slide numeric contradictions); Track Changes Ctrl+Alt+K (vs SharePoint history or file, accept/decline); change language; Finalize (strip links, refresh all, PDF/PPTX of deck or selection; options: break Excel links, remove notes/comments, PDF appendix insertion); Send (finalize + email). PPT-side ELM Ctrl+Shift+R; update linked item Ctrl+Shift+E; Quick Actions Ctrl+Space; send selection Ctrl+Shift+M; AI Assistant Ctrl+Shift+Space (condense/rephrase/translate FR/EN/DE); AI Glossary (auto abbreviation glossary slide); Claude for PPT/Excel integration; Gantt Charts; Slide Converter (rebrand whole deck); Facing Pages (duplex parts).

## Word
Dedicated ribbon; Templates; Library; Smart Check Ctrl+Shift+C (double spaces, punctuation spacing by language, number-format separators); number-to-words Ctrl+Shift+N; Excel linking (export Ctrl+Shift+D, update Ctrl+Shift+U, Update All pane).

## Power BI
PBI->PPT/Word link (visuals or pages as bitmap via WebView2; needs Pro/Premium + cloud-hosted reports); import panel with filters/slicers before export; refresh data; Update Links panel (title, author, modified, source, filters); open source report.

## Admin / deployment / governance (titles-level)
Smart Format theme management; PPT/Word template managers; breadcrumb/template wiring; library structure + tags; Proposal Wizard editing; TOC automation setup (8-part); per-group content sharing; change notifications + user suggestions; pre-linked library content; Dynamic Library data sources; license portal (seats, CSV import, pending requests, roles); shared path/team settings; user-group settings automation; SSO; Entra ID provisioning; .exe/SCCM/Intune/Citrix/web-add-in deployment; update + monitoring services; install validation; .UpSlide config bundles; library hosting (SharePoint/Azure Blob/shared folder); PBI link consent; Outlook Signature Manager; Gantt rollout; security/architecture docs; UpSlide MCP (+ prepare library for MCP, connect Claude); AI custom actions, model governance, privacy, AI credits; iManage compatibility.

## Coverage gaps
Admin articles mostly title-level; release notes, FAQs, tutorials, troubleshooting categories unexplored; Word deep-dives thin; PBI FAQ unopened; non-English content not surveyed. Full fetched-URL list lives in the agent transcript.
