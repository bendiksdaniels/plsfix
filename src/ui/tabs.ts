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

  for (const tab of tabs) {
    tab.addEventListener("click", () => activate(tab.id));
  }

  return { activate };
}
