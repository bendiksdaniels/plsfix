# AppSource listing copy for pls,fix

Everything the Partner Center forms ask for, ready to paste. Steps and field names follow
Microsoft's submission guide (`docs/appsource/checklist.md`). English only for the first
submission; the manifest's DefaultLocale is en-US.

## Offer

- **Offer name:** pls,fix (identical to the manifest's DisplayName)
- **Publisher:** Daniels Bendiks (must equal the manifest's ProviderName; immutable)
- **Product setup:** not in the Apple Store; no Microsoft Entra ID / SSO; no additional
  purchases; no CRM connection
- **Package:** `manifest.prod.xml` from the latest release (store validation must pass first,
  see the checklist)

## Properties

- **Categories (up to 3):** Productivity; Data analytics; Collaboration
- **Industries (up to 2):** Financial services; Professional services
- **EULA:** Microsoft's Standard Contract (accept in the dialog)
- **Privacy policy link:** https://dbautomatizacijas.com/modelis/privacy.html
- **Support document link:** https://dbautomatizacijas.com/modelis/support.html

## Marketplace listing (English)

**Summary** (the one-liner under the name):

> Format, audit and link financial models: consistent formatting cycles, model checks, and
> Excel tables and charts that stay fresh in PowerPoint.

**Description** (HTML allowed; paste as is):

pls,fix is a free add-in for analysts who build financial models. It gives Excel the tools
the COM add-ins of the big houses are known for, as a web add-in that runs on Windows, Mac
and Excel on the web, and it keeps a PowerPoint deck in sync with the model without
re-pasting.

**Model formatting.** Five brand presets, number-format cycles on the keys you already know
(Ctrl+Shift+1 general, +4 currency, +5 percent), title, result and item row styles, fill and
font colour cycles, borders, row heights and column widths. A paintbrush with three slots that
travel with the workbook. Autocolor: hardcodes blue, formulas black, links green, with a
colour key on demand.

**Auditing.** A formula-consistency overlay that stripes formulas matching their neighbours
and reddens the ones that break the pattern, Smart Track for precedents and dependents, a
model check that lists errors, hardcodes inside formulas, inconsistent formulas, volatile
functions, broken names, unused styles, hidden sheets and external links with a jump to each
cell, Super Find across every sheet, and Prepare for sharing.

**Formulas and charts.** Fill formulas sized by the neighbouring rows, paste values, formats,
exact formulas or transposed, reversible IFERROR, CAGR, sign flip, x1000, six ready
calculation blocks (debt schedule, DCF, NPV/IRR, working-capital days, sensitivity grid,
EBITDA bridge), a branded waterfall from a bridge table, a tornado chart, unpivot. Consistent
rounding through the custom functions =PLSFIX.ROUND and =PLSFIX.ROUNDSUM: the rounded parts
add up to the rounded total.

**Linked objects in PowerPoint.** Export a range as a picture or an editable table, a cell as
a text box, or a chart as native shapes, and update the deck with one click when the model
changes. Position and size are kept. Links are encrypted on your device before they travel;
the relay stores ciphertext only.

**Object tools in PowerPoint.** Align, distribute, match size, select similar, swap, and a
Smart Painter for fills and outlines, on the ribbon and in the pane.

Every section of the pane carries a "?" that explains its buttons, and the Tools tab prints a
keyboard shortcut card. No account, no sign-in, no telemetry. Open source under the MIT
license: https://github.com/bendiksdaniels/plsfix

**Search keywords (3):** financial modelling; Excel formatting; Excel to PowerPoint links

**Icons:** the listing logo is the add-in icon at the sizes the form asks for
(`public/assets/icon-32.png`, `icon-64.png`, `icon-80.png` scaled from the same artwork;
check the form's pixel sizes when uploading).

**Screenshots (1366 x 768, at least one):** see `docs/appsource/screenshots/` once captured
(the Excel pane on the demo model with the audit overlay on; the PowerPoint Links tab with an
inserted table; the Tools tab with the templates). Captions:

1. Autocolor and the audit overlay on a P&L: hardcodes, formulas and the one broken formula.
2. A table and a chart from the model, inserted in PowerPoint and updated in place.
3. Format cycles, paintbrush slots and the six calculation blocks in the Tools tab.

## Availability

All markets, available as soon as approved (the schedule cannot be changed after the first
publish). Free.

## Notes for certification

Paste `docs/appsource/test-notes.md` into the **Notes for certification** box (it includes the
custom-function test the store requires).
