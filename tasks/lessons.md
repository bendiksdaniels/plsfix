# Lessons

## 2026-08-27: the `hidden` attribute is defeated by any author `display` rule

`.view { display: grid }` silently killed `element.hidden = true` (UA-origin `[hidden]{display:none}`
loses to any author-origin display declaration, regardless of specificity). Tab switching shipped
broken; tests and build were green because neither sees CSS cascade.

Rules:
- Every stylesheet in this repo carries `[hidden] { display: none !important; }` in the reset.
- A UI feature is not "done" on tests+build alone: render the pane (browser dev mode or sideload)
  and click the new interaction before claiming completion. The shared Playwright browser can be
  held by another session — fall back to manual sideload confirmation and say so.
- When hiding a container of interactive children, also clear its children (stale listeners on
  visible-again nodes re-apply old state).
