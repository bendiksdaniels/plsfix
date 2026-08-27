# User-reported gaps and pains in modelling add-ins (research agent, 2026-08-27)

Sources: Capterra/SoftwareAdvice reviews (fetched), G2 (search-derived, fetch-blocked), WSO forums,
vendor docs, think-cell KB. Reddit/MrExcel/Quora unreachable (403s) — noted as a gap, not absence.

## Ranked gap list (most frequent first)

| # | Gap/pain | Evidence | Freq | Implication for our v1 |
|---|---|---|---|---|
| 1 | Price disproportionate for individuals/small teams | MB Capterra "more affordable pricing for individual users", value 2/5; US G2 "licensing costs too expensive to provide every employee"; TC "affordable for corporate users only" ($327.60/user/yr entry) | 6+ | A credible lean/individual tier is an unmet want across all incumbents. Business decision for Daniel. |
| 2 | Performance degrades on large files / after repeated actions | MB "when files are huge... the model will hang", "almost unusable" in large workbooks, Excel startup delay; TC "slow after each action"; US "library slow during updates" | 5 | Benchmark on deal-sized models; keep batch-sync discipline, caps, untrack; never block the pane. |
| 3 | Conflicts with other add-ins, above all Capital IQ | MB reviewers x4: "interference with the Cap IQ plugin", "incapable of use with S&P Capital IQ", freezing | 4+ | COM-era pain we mostly sidestep (web add-in), but shortcut-conflict UX matters (Office shows a picker); test alongside CapIQ/FactSet ribbons. |
| 4 | Learning curve / feature discoverability / thin docs | MB "learning process is long", "underutilized functions with inefficient discovery"; TC "not able to know about its features from its interface"; US "many features I will likely never use" | 5 | In-product discoverability: kbd hints on buttons (have), shortcut cheat card, sane defaults, short README GIFs. |
| 5 | Undo unreliable / destructive bulk actions | MB reviewers x3 (all 5/5 overall!): "Can't undo with the hot keys", "actions cannot be undone if used by mistake", "Undo functionality is very scary" | 3+ | SEVERITY-1 CLASS. Office.js writes also bypass native Ctrl+Z. Build "SMT Undo": snapshot affected range state before every bulk mutation, one-tap restore of the last action. Differentiator. |
| 6 | IT/deployment fragility: breaks after Windows/Office updates, settings don't roam, silent disable | MB "after Windows updates, must repeatedly reenable or reinstall", "settings don't persist across computers", "plug-in disappearing" | 4 | Web add-in + centralized deployment sidesteps the COM auto-disable class. Settings roaming: publishable settings (M4) or file/Graph-backed. |
| 7 | Excel->PPT refresh is one link at a time | MB "helpful if it allowed updating all at once"; US library slides update individually | 2 | One-click "refresh everything in this deck" is a stated missing feature — M2/M3 requirement, keep Update All first-class. |
| 8 | Exports lose editability / break for non-licensed recipients | US "charts imported as images... difficult to edit"; TC "when a non-user edits a think-cell chart all functionality is lost" | 2 | Exports must degrade gracefully without the add-in (image + intact native objects; never proprietary-only). |
| 9 | Named misses: VBA hooks, custom shortcut builder, library subfolders, more chart types, column-width formula | MB/US/TC scattered singles | 4-5 thin | Custom-shortcut layer (replaceShortcuts) is our cheap win; others note-only. |
| 10 | Free/DIY (QAT + dead macro packs) tops out fast | WSO: TTS Turbo dead since 2019, tracer "not always stable"; "no free macros as good as the paid ones" | 2+ | A genuinely good free tier is a wedge if it clears the QAT bar (ours already does). |

## Mac support (the wedge)

- UpSlide own docs: "UpSlide is a VSTO add-in and cannot be installed on a Mac." Workaround = Windows VM.
- Macabacus own docs: "Office for Mac does not support sophisticated COM add-ins like Macabacus, and Microsoft is not expected to change this." Mac users get feature-reduced "Macabacus Lite"; full product needs Windows.
- think-cell needed years to port; Apple Silicon still has launch friction (KB0238).
- Few direct "wish it ran on Mac" reviews = selection bias (banking is Windows-issued); think-cell's consultant/corporate audience shows the visible complaint history.
- IMPLICATION: our Office.js add-in is natively cross-platform (Windows/Mac/web). Lead the positioning with it for corporate finance, boutique advisory, consulting-adjacent and Mac-based analysts.

## Pricing sentiment

- Macabacus ~$200+/yr individual, ~$1,000/yr per 5 seats; personal tier chatter $10-20/mo; 4.7/5 overall so target buyers tolerate it; individuals priced out.
- UpSlide enterprise quote-only; "maybe the most expensive" (one 3-star review).
- think-cell $327.60/user/yr entry scaling to ~$252/seat at 50; accepted premium for power users, occasional users priced out.
- Category survey (UpSlide CEO, ~600 bank IT leaders): two-thirds say >25% of software spend wasted on unused tools.
- Pattern: heavy users accept premium; the gap is individuals/small teams/occasional users. Segmentation signal.

## What users love (table stakes — do not cut)

- Shortcut-driven formatting that replaces repetitive work (MB).
- Formula auditing / precedents-dependents tracing on inherited messy models (WSO: "great for when you get a piece of bad model").
- Live Excel->PPT/Word linking with one-click refresh ("clutch when refreshing numbers").
- Presentation chart types PowerPoint lacks: waterfall, Gantt, Marimekko (think-cell's whole business).
- Responsive support (US 9.5 vs MB 8.1 on G2 support quality).
- Governed brand/template library at enterprise scale.
- Low-friction free on-ramp (Macabacus Lite cited as the entry drug).

## Coverage gaps (from the agent)

Reddit entirely unreachable (403 walls, ~2,050 indexed threads exist — real gap); G2/TrustRadius quotes are search-derived, not direct reads; MrExcel/Quora 403; HN has literally zero results for these products (not a general-tech tool); CapIQ plug-in has thin standalone review coverage; UpSlide pricing unconfirmable (quote-only).
