# Competitor delta features beyond Macabacus/UpSlide (research agent, 2026-08-27)

Deltas only — features the MB/US inventories do not cover. Snippet-sourced rows flagged; bpmglobal.com refused all connections.

## think-cell
- Consistent rounding engine (TCROUND/TCROUNDUP/TCROUNDDOWN): picks which values to round against the grain so rounded subtotals/totals sum exactly; #NUM! when impossible. Kills "rounded table doesn't add up". (think-cell.com/en/resources/manual/excel-data-rounding)
- Calendar-native Gantt with dependency-aware relayout (product/gantt)
- Deck sanitize before external distribution (hidden data, notes, metadata) (manual/moretools)
- Global decimal-symbol switch (comma vs period locales, whole deck) (manual/moretools)

## S&P Capital IQ Excel plug-in
- 250+ ready-made model templates instantly populated by identifier ("New From Template" -> populated tearsheet/comps/credit workbook). (WU library PDF, pp.1-10)
- Multilingual template outputs (JA/KO/zh-CN). (same)
- Excluded as repeats: 150+ Quick Keys, Formula Builder.

## FactSet for Office
- Filing Wizard: section-level filing extraction into Excel. (snippet; no Macabacus-style formatting suite found — flagged, not padded)

## Arixcel Explorer
- Row/column-insertion-aware workbook diff: re-aligns sheets before diffing so one inserted row doesn't flag the whole sheet. (snippet)
- Calculation Flow map: classifies cells input/calc/output, shows data-flow direction. (snippet)

## Operis Analysis Kit (OAK)
- Logic-aware circular reference detection: resolves active IF branches, finds latent circularity Excel's warning misses. (operisanalysiskit.com/modelling-insights/how-to-find-circular-references-in-excel/)
- Model complexity/risk score + review issues list. (operisanalysiskit.com)

## Modano
- Modular reusable model components (central "Lego" library across workbooks)
- ERP/accounting-linked month-end roll-forward (auto-extend with actuals)
- Schema-level structural automation (new product line propagates across sheets)
(modano.com/solutions/financial-modeling)

## F1F9 bpmToolbox (ALL snippet-sourced; site refused connections)
- Base sheets (new sheets seed from standardized template)
- Keys: auto-maintained color/format legend in sync with the model
- Automated time-series roll-forward
- Multi-workbook batch restructuring

## PowerUser
- Data maps (350+ choropleth maps linked to Excel), Sankey/Tornado/Dumbbell charts, Tombstones, deck-wide Draft/Confidential Stamps, UnPivot, Smart Fields (deck-wide text tokens), Vectorize Text, Select Similar Shapes. (powerusersoftwares.com/features)

## QuickCel
- Multi Goal Seek (several targets at once); everything else repeats Macabacus. (quickcel.software)

## Ablebits Ultimate Suite
- Fuzzy duplicate detection; merge tables by key columns; Unpivot; regex find/replace/extract; random data generator; Spell Number (numbers to words w/ currency); Formula Editor with tree view. (ablebits.com/excel-suite)

## Kutools
- Super Lookup suite (multi-condition, one-to-many); live currency conversion; "Make Up a Number" subset-sum solver (which cells sum to a target — variance reconciliation!); Number to Words; Fuzzy Lookup; Error Condition Wizard (custom error messages). (extendoffice.com feature list)

## empower Suite
- One-click conversion of rival add-ins' charts to native format (migration play)
- Move-resilient Excel links (relative-path logic survives moved files)
(empowersuite.com; other features repeat MB/US)

# Cross-product patterns worth copying
- Unpivot/flatten (PowerUser + Ablebits)
- Fuzzy text matching (Ablebits + Kutools)
- Number-to-words (Ablebits + Kutools; also UpSlide Word)
- Automated period roll-forward (Modano + bpmToolbox)
- Gantt (think-cell + PowerUser + empower)
- Two-version workbook diff (Arixcel + OAK + Ablebits + Kutools) — version-to-version review, distinct from in-file auditing
- Standardized building blocks (Modano modules + bpmToolbox base sheets)

# Coverage gaps
Arixcel Accounts unverifiable; FactSet formatting suite absent from public docs; bpmToolbox primary sources unreachable; roundups surfaced mostly generic doc-automation SaaS (excluded); no hands-on testing — all vendor-doc/snippet sourced.
