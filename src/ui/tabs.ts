// Tab switcher: reads [role=tab] buttons under a bar, each pointing at a
// panel via aria-controls. Clicking a tab (or calling activate) shows that
// panel and hides the rest. It never touches panel contents - a stale
// listener inside a panel is the caller's problem to solve on re-render.
// Invariant: the active tab is the strip's one Tab stop, and every way of
// activating a tab goes through its own click, so a caller's hook fires.

export function installTabs(bar: HTMLElement): {
  activate(tabId: string): void;
} {
  const tabs = Array.from(bar.querySelectorAll<HTMLElement>("[role=tab]"));

  function activate(tabId: string): void {
    for (const tab of tabs) {
      const isActive = tab.id === tabId;
      tab.classList.toggle("active", isActive);
      tab.setAttribute("aria-selected", String(isActive));
      // Roving tabindex, the ARIA tabs pattern: one press of Tab enters the
      // strip and the next leaves it, however many tabs the pane has, and the
      // arrows below move between them.
      tab.tabIndex = isActive ? 0 : -1;

      const panelId = tab.getAttribute("aria-controls");
      const panel = panelId ? document.getElementById(panelId) : null;
      if (panel) panel.hidden = !isActive;
    }
  }

  // The ARIA tabs pattern: Left/Right (Up/Down too, in case a host ever
  // stacks the bar) move focus AND activate in one step, Home/End jump to
  // the ends, and the move wraps past either edge. Tab itself is left alone
  // so it still leaves the strip on the first press - nothing here traps it.
  // click(), not activate(): the panes hang their own work on a tab's click
  // (the Excel Workbook tab re-reads the sheet explorer there), and a keyboard
  // user has to get the tab a mouse user gets, not a quieter one.
  function moveTo(index: number): void {
    const tab = tabs[(index + tabs.length) % tabs.length];
    if (!tab) return;
    tab.focus();
    tab.click();
  }

  function onKeydown(event: KeyboardEvent): void {
    const current = tabs.findIndex((tab) => tab === event.target);
    if (current === -1) return;
    switch (event.key) {
      case "ArrowRight":
      case "ArrowDown":
        event.preventDefault();
        moveTo(current + 1);
        break;
      case "ArrowLeft":
      case "ArrowUp":
        event.preventDefault();
        moveTo(current - 1);
        break;
      case "Home":
        event.preventDefault();
        moveTo(0);
        break;
      case "End":
        event.preventDefault();
        moveTo(tabs.length - 1);
        break;
      default:
        break;
    }
  }

  for (const tab of tabs) {
    tab.addEventListener("click", () => activate(tab.id));
    tab.addEventListener("keydown", onKeydown);
  }

  // The markup ships one tab selected; the roving tabindex has to match it
  // before the first activation, or every tab starts out a Tab stop. Only the
  // tabindex is set here - which panel is showing stays the markup's word.
  const open = tabs.find((tab) => tab.getAttribute("aria-selected") === "true");
  for (const tab of tabs) tab.tabIndex = tab === (open ?? tabs[0]) ? 0 : -1;

  return { activate };
}
