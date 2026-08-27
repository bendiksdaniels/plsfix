# Product roadmap

## Milestone 1 — Excel modelling core

- [x] Task-pane shell and local sideload manifest
- [x] Selection inspector
- [x] Branded cell presets and number formats
- [x] Formula fill, IFERROR, scaling, and autocolor
- [x] User-editable brand palette and font settings
- [ ] Ribbon tab and keyboard shortcuts (shared runtime)
- [ ] Palette-driven formatting cycles
- [ ] Autocolor v2: external links, partial inputs, legend
- [ ] Formula consistency overlay
- [ ] Fast fill auto-extent and paste-special suite
- [ ] Quick CAGR, sign flip, decimal steppers
- [ ] Sheet explorer pane and name scrubber
- [ ] Workbook table of contents
- [ ] Precedent/dependent navigation
- [ ] Waterfall chart builder

## Milestone 2 — Linking proof of concept

- [ ] PowerPoint companion manifest and task pane
- [ ] Export an Excel range as a high-resolution image
- [ ] Persist workbook, worksheet, and range link metadata
- [ ] Choose the link relay: Rust service or Microsoft Graph
- [ ] Update one PowerPoint object from its Excel source
- [ ] Detect missing and ambiguous sources

## Milestone 3 — Link manager

- [ ] List and filter all links in a presentation
- [ ] Update selected, slide, or all links
- [ ] Change source and resolve workbook versions
- [ ] Highlight linked cells in Excel
- [ ] Preserve position and size during refresh
- [ ] Performance and failure-isolation testing

## Milestone 4 — Enterprise product

- [ ] Microsoft Entra ID authentication
- [ ] Tenant-level configuration and feature flags
- [ ] Shared settings publishing and shortcut manager
- [ ] Shared template and content library
- [ ] Admin deployment package
- [ ] Telemetry with explicit privacy controls
- [ ] Windows, Mac, and web compatibility matrix
