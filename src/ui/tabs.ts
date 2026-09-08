// Tab switcher: reads [role=tab] buttons under a bar, each pointing at a
// panel via aria-controls. Clicking a tab (or calling activate) shows that
// panel and hides the rest. It never touches panel contents - a stale
// listener inside a panel is the caller's problem to solve on re-render.

export function installTabs(bar: HTMLElement): {
  activate(tabId: string): void;
} {
  const tabs = Array.from(bar.querySelectorAll<HTMLElement>("[role=tab]"));

  function activate(tabId: string): void {
    for (const tab of tabs) {
      const isActive = tab.id === tabId;
      tab.classList.toggle("active", isActive);
      tab.setAttribute("aria-selected", String(isActive));

      const panelId = tab.getAttribute("aria-controls");
      const panel = panelId ? document.getElementById(panelId) : null;
      if (panel) panel.hidden = !isActive;
    }
  }

  // The ARIA tabs pattern: Left/Right (Up/Down too, in case a host ever
  // stacks the bar) move focus AND activate in one step, Home/End jump to
  // the ends, and the move wraps past either edge. Tab itself is left alone
  // so it still leaves the strip on the first press - nothing here traps it.
  function moveTo(index: number): void {
    const tab = tabs[(index + tabs.length) % tabs.length];
    if (!tab) return;
    tab.focus();
    activate(tab.id);
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

  return { activate };
}
