// The in-page half of the UX gate: injected as a classic script and called
// back through window.__ux.collect(). It owns every measurement taken inside
// the pane - horizontal overflow, the tab strip, tap-target heights, clipped
// button labels, tiny text and a toast sitting over the tabs - and returns
// plain strings, because nothing but strings survives the bridge back.

(function installUxProbe() {
  var EPS = 0.5;
  var MIN_TAP_HEIGHT = 32;
  var MIN_FONT_PX = 11;

  function describe(el) {
    if (el.id) return "#" + el.id;
    var s = el.tagName.toLowerCase();
    if (typeof el.className === "string" && el.className.trim()) {
      s += "." + el.className.trim().split(/\s+/).join(".");
    }
    var parent = el.parentElement;
    if (parent) {
      s +=
        "[" + String(Array.prototype.indexOf.call(parent.children, el)) + "]";
    }
    return s;
  }

  function isRendered(el) {
    var rects = el.getClientRects();
    if (rects.length === 0) return false;
    if (rects[0].width === 0 && rects[0].height === 0) return false;
    var cs = getComputedStyle(el);
    return cs.visibility !== "hidden" && cs.display !== "none";
  }

  function hasOwnText(el) {
    for (var node of el.childNodes) {
      if (node.nodeType === 3 && (node.textContent || "").trim().length > 0) {
        return true;
      }
    }
    return false;
  }

  function isTapControl(el) {
    var type = (el.getAttribute("type") || "").toLowerCase();
    if (el.tagName === "BUTTON" || el.tagName === "SELECT") return true;
    var exempt = ["checkbox", "color", "file", "hidden"];
    return el.tagName === "INPUT" && exempt.indexOf(type) === -1;
  }

  // A button label that does not fit its button is a bug however it fails -
  // ellipsised, or bleeding past the edge into the next control. Deliberately
  // truncated *data* (a workbook name, a cell address) is the design, so this
  // only looks at buttons and at a column header, which is a fixed string the
  // column has to be wide enough for.
  function isClippedLabel(el) {
    if (el.scrollWidth <= el.clientWidth + EPS) return false;
    if (el.tagName === "TH") return true;
    return el.tagName === "BUTTON" || el.closest("button") !== null;
  }

  function scanElement(el, vw, out) {
    var rect = el.getBoundingClientRect();
    if (rect.right > vw + EPS) {
      out.overflow.push(
        describe(el) + " right=" + rect.right.toFixed(1) + " viewport=" + vw,
      );
    }
    if (
      isTapControl(el) &&
      rect.height > 0 &&
      rect.height < MIN_TAP_HEIGHT - EPS
    ) {
      out.shortControls.push(
        describe(el) + " height=" + rect.height.toFixed(1),
      );
    }
    if (isClippedLabel(el)) {
      out.clippedLabels.push(
        describe(el) +
          " scrollWidth=" +
          String(el.scrollWidth) +
          " clientWidth=" +
          String(el.clientWidth) +
          ' text="' +
          (el.textContent || "").trim().slice(0, 60) +
          '"',
      );
    }
    // No grace band on font-size: it is an authored exact value, not a layout
    // rounding artefact. The 0.05 guard is for DPI float noise only.
    if (hasOwnText(el)) {
      var fontPx = parseFloat(getComputedStyle(el).fontSize);
      if (fontPx < MIN_FONT_PX - 0.05) {
        out.smallText.push(
          describe(el) +
            " font=" +
            fontPx +
            'px text="' +
            (el.textContent || "").trim().slice(0, 40) +
            '"',
        );
      }
    }
  }

  // Every band that owns a width of its own: a section that scrolls sideways
  // is the same defect as a body that does, one level down.
  function scanScrollers(out) {
    var boxes = document.querySelectorAll(
      "body, main, .view, section, .tab-bar",
    );
    for (var el of boxes) {
      if (!isRendered(el)) continue;
      if (el.scrollWidth > el.clientWidth + EPS) {
        out.scrollers.push(
          describe(el) +
            " scrollWidth=" +
            String(el.scrollWidth) +
            " clientWidth=" +
            String(el.clientWidth),
        );
      }
    }
  }

  function scanTabStrip(out) {
    var bar = document.getElementById("tab-bar");
    if (!bar || !isRendered(bar)) return;
    var barRect = bar.getBoundingClientRect();
    var tabs = bar.querySelectorAll("[role=tab]");
    for (var tab of tabs) {
      var rect = tab.getBoundingClientRect();
      if (rect.left < barRect.left - EPS || rect.right > barRect.right + EPS) {
        out.tabStrip.push(
          describe(tab) +
            " sticks out of the bar (tab " +
            rect.left.toFixed(1) +
            "-" +
            rect.right.toFixed(1) +
            ", bar " +
            barRect.left.toFixed(1) +
            "-" +
            barRect.right.toFixed(1) +
            ")",
        );
      }
      if (tab.scrollWidth > tab.clientWidth + EPS) {
        out.tabStrip.push(
          describe(tab) +
            ' label clipped: "' +
            (tab.textContent || "").trim() +
            '" needs ' +
            String(tab.scrollWidth) +
            "px, has " +
            String(tab.clientWidth) +
            "px",
        );
      }
    }
  }

  function overlaps(a, b) {
    return (
      a.left < b.right - EPS &&
      a.right > b.left + EPS &&
      a.top < b.bottom - EPS &&
      a.bottom > b.top + EPS
    );
  }

  function scanToast(out) {
    var toast = document.getElementById("toast");
    var bar = document.getElementById("tab-bar");
    if (!toast || !bar) return;
    if (!toast.classList.contains("visible")) return;
    if (overlaps(toast.getBoundingClientRect(), bar.getBoundingClientRect())) {
      out.toast.push("#toast covers #tab-bar");
    }
  }

  function collect() {
    var vw = window.innerWidth;
    var out = {
      overflow: [],
      scrollers: [],
      tabStrip: [],
      shortControls: [],
      clippedLabels: [],
      smallText: [],
      toast: [],
      viewportWidth: vw,
    };
    for (var el of document.querySelectorAll("body *")) {
      if (isRendered(el)) scanElement(el, vw, out);
    }
    scanScrollers(out);
    scanTabStrip(out);
    scanToast(out);
    return out;
  }

  window.__ux = { collect: collect };
})();
