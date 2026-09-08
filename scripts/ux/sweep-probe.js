// The in-page half of the click sweep (scripts/ux/click-sweep.mjs). Owns:
// error capture (window.onerror/unhandledrejection, the same events
// src/ui/report.ts listens for), the target catalogue, and one click's
// before/after comparison. A "reaction" is any of: document.body.innerHTML
// changed, any input/select/textarea's value or checked state changed, the
// document's active element changed, or document.title changed - between
// them these catch a toast, a re-rendered list, a hidden/aria/class flip and
// a focus move without needing to enumerate every attribute a fix might use.

(function installSweepProbe() {
  var errors = [];

  function installHooks() {
    window.addEventListener("error", function (event) {
      errors.push("error: " + describeThrown(event.error ?? event.message));
    });
    window.addEventListener("unhandledrejection", function (event) {
      errors.push("unhandledrejection: " + describeThrown(event.reason));
    });
  }

  function describeThrown(value) {
    if (value instanceof Error) return value.message;
    try {
      return String(value);
    } catch {
      return "(unprintable error)";
    }
  }

  function controlsFingerprint() {
    var nodes = document.querySelectorAll("input, select, textarea");
    var parts = [];
    for (var el of nodes) {
      if (el.type === "checkbox" || el.type === "radio") {
        parts.push(el.checked ? "1" : "0");
      } else {
        parts.push(el.value);
      }
    }
    return parts.join("|");
  }

  function snapshot() {
    return {
      html: document.body.innerHTML,
      title: document.title,
      controls: controlsFingerprint(),
      activeId: document.activeElement ? document.activeElement.id : "",
      errorCount: errors.length,
    };
  }

  function reacted(before, after) {
    return (
      before.html !== after.html ||
      before.title !== after.title ||
      before.controls !== after.controls ||
      before.activeId !== after.activeId
    );
  }

  // Polls every POLL_MS up to budgetMs for a reaction or a new error, so the
  // common (fast) case returns quickly instead of always paying the ceiling.
  var POLL_MS = 25;
  async function waitForReaction(before, budgetMs) {
    var waited = 0;
    while (waited <= budgetMs) {
      var after = snapshot();
      if (after.errorCount > before.errorCount) {
        return { reacted: false, erred: true };
      }
      if (reacted(before, after)) return { reacted: true, erred: false };
      await new Promise(function (resolve) {
        setTimeout(resolve, POLL_MS);
      });
      waited += POLL_MS;
    }
    return { reacted: false, erred: false };
  }

  // A target inside a tab panel that is not the active one is switched to
  // first, with a real click on its own tab - the same way a modeller would
  // actually reach it, and the only way to prove the panel itself reacts.
  function surfaceOwnTab(el) {
    var panel = el.closest('[role="tabpanel"]');
    if (!panel || !panel.hidden) return;
    var tab = document.querySelector(
      '[role="tab"][aria-controls="' + panel.id + '"]',
    );
    if (tab) tab.click();
  }

  // Fires a real DOM click and reports what happened within budgetMs. kind
  // "checkbox" toggles + fires change; "select" moves to the next option
  // (wrapping) + fires change - with only zero or one option (nothing
  // linked, no host to populate it) there is nothing to change, so this is
  // skipped like a disabled button rather than failed as "no reaction";
  // anything else calls element.click(). An already-active tab is clicked
  // (proving it stays consistent) but not held to producing a reaction -
  // re-selecting the current tab is correctly a no-op.
  async function clickAndObserve(selector, kind, budgetMs) {
    var el = document.querySelector(selector);
    if (!el) return { found: false };
    surfaceOwnTab(el);
    if (el instanceof HTMLButtonElement && el.disabled) {
      return { found: true, skippedDisabled: true };
    }
    if (kind === "select" && el.options.length <= 1) {
      return { found: true, skippedNoOptions: true };
    }
    var alreadyActiveTab =
      kind === "click" &&
      el.getAttribute("role") === "tab" &&
      el.getAttribute("aria-selected") === "true";
    var before = snapshot();
    if (kind === "checkbox") {
      el.checked = !el.checked;
      el.dispatchEvent(new Event("change", { bubbles: true }));
    } else if (kind === "select") {
      el.selectedIndex = (el.selectedIndex + 1) % el.options.length;
      el.dispatchEvent(new Event("change", { bubbles: true }));
    } else {
      el.click();
    }
    var result = await waitForReaction(before, budgetMs);
    if (alreadyActiveTab && !result.reacted && !result.erred) {
      result.reacted = true;
      result.skippedAlreadyActive = true;
    }
    return { found: true, skippedDisabled: false, ...result };
  }

  // Every [data-action] button plus every other id the sweep drives: tabs,
  // help toggles, first-run buttons, selects and checkboxes. Excludes
  // buttons a real click-sweep should not press blind (the styles/names
  // delete confirms only arm on their first click, which is exercised
  // through the click, but a second sweep pass never runs to avoid a real
  // delete call reaching a mocked adapter twice in the same state).
  function listTargets() {
    var out = [];
    for (var button of document.querySelectorAll("button[data-action]")) {
      out.push({
        kind: "button",
        selector: cssPathFor(button),
        label: textOf(button),
        action: button.dataset.action,
      });
    }
    for (var tab of document.querySelectorAll('[role="tab"]')) {
      out.push({ kind: "tab", selector: "#" + tab.id, label: textOf(tab) });
    }
    for (var toggle of document.querySelectorAll(".help-toggle")) {
      out.push({
        kind: "help",
        selector: "#" + toggle.id,
        label: toggle.getAttribute("aria-label") || toggle.id,
      });
    }
    // Buttons the first-run card wires by id, not data-action (its Shortcut
    // card button is the "shortcut card link" the sweep is asked to press
    // directly, on top of the data-action button that opens the same card).
    // installHelp() also hangs a "?" toggle on this same section, which the
    // .help-toggle pass above already catches - skipped here so it is not
    // clicked (and its open/close state disturbed) a second, redundant time.
    for (var frButton of document.querySelectorAll(
      ".first-run button:not(.help-toggle)",
    )) {
      out.push({
        kind: "button",
        selector: "#" + frButton.id,
        label: textOf(frButton),
        action: frButton.id,
      });
    }
    for (var check of document.querySelectorAll('input[type="checkbox"]')) {
      out.push({
        kind: "checkbox",
        selector: idSelector(check),
        label: labelFor(check),
      });
    }
    for (var select of document.querySelectorAll("select")) {
      out.push({
        kind: "select",
        selector: idSelector(select),
        label: labelFor(select),
      });
    }
    return out;
  }

  function idSelector(el) {
    if (!el.id)
      throw new Error(
        "sweep-probe: control with no id: " + el.outerHTML.slice(0, 80),
      );
    return "#" + el.id;
  }

  function labelFor(el) {
    var aria = el.getAttribute("aria-label");
    if (aria) return aria;
    var row = el.closest(".io-row, .field, .chip, div");
    var label = row ? row.querySelector("label") : null;
    return (label ? label.textContent : el.id) || el.id;
  }

  function textOf(el) {
    return (el.textContent || "").replace(/\s+/g, " ").trim();
  }

  // Only used for a data-action button that also lacks an id: data-action
  // values are unique by construction (dispatch() switches on them), so the
  // attribute selector alone is already a stable, human-readable target.
  function cssPathFor(button) {
    if (button.id) return "#" + button.id;
    return '[data-action="' + button.dataset.action + '"]';
  }

  function activeTabId() {
    var el = document.querySelector('[role="tab"][aria-selected="true"]');
    return el ? el.id : null;
  }

  // Exactly one panel visible, and it is the active tab's own one.
  function tabState() {
    var visible = [];
    for (var tab of document.querySelectorAll('[role="tab"]')) {
      var panelId = tab.getAttribute("aria-controls");
      var panel = panelId ? document.getElementById(panelId) : null;
      if (panel && !panel.hidden) visible.push(panelId);
    }
    return { visiblePanels: visible, activeTabId: activeTabId() };
  }

  function openHelpCards() {
    var open = [];
    for (var card of document.querySelectorAll(".help-card")) {
      if (!card.hidden) open.push(card.id);
    }
    return open;
  }

  // Two buttons sharing one data-action (or one first-run id) would let a
  // search row silently click the wrong one - findButton() in
  // src/pane/tool-search.ts returns the first DOM match, not necessarily
  // the button the row was built from.
  function duplicateActionKeys() {
    var seen = {};
    var dupes = [];
    var buttons = document.querySelectorAll(
      "button[data-action], .first-run button[id]",
    );
    for (var button of buttons) {
      var key = button.dataset.action || button.id;
      if (seen[key]) dupes.push(key);
      seen[key] = true;
    }
    return dupes;
  }

  // Types a tool's own label, clicks its top search result and reports
  // whether the pane landed on the tab the row itself claimed.
  async function checkSearchRow(label) {
    var input = document.getElementById("tool-search");
    var list = document.getElementById("tool-search-results");
    if (!input || !list) return { skipped: true };
    input.value = label;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    await new Promise(function (resolve) {
      setTimeout(resolve, 30);
    });
    var row = list.querySelector("li");
    if (!row) return { found: false };
    var badge = row.querySelector(".tool-search-tab");
    var rowTab = badge ? textOf(badge) : "";
    row.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    await new Promise(function (resolve) {
      setTimeout(resolve, 30);
    });
    var landedId = activeTabId();
    var landedTab = landedId ? textOf(document.getElementById(landedId)) : "";
    return {
      found: true,
      rowTab: rowTab,
      landedTab: landedTab,
      ok: rowTab === "" || rowTab === landedTab,
    };
  }

  window.__sweep = {
    installHooks: installHooks,
    snapshot: snapshot,
    listTargets: listTargets,
    clickAndObserve: clickAndObserve,
    tabState: tabState,
    openHelpCards: openHelpCards,
    duplicateActionKeys: duplicateActionKeys,
    checkSearchRow: checkSearchRow,
    errors: function () {
      return errors.slice();
    },
  };
})();
