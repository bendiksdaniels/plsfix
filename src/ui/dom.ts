// The one element lookup both panes use. A pane is wired to its own HTML, so a
// missing id is a build mistake, not a runtime condition: it throws by name
// rather than handing back a null every caller would have to check.

export function getElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing element #${id}`);
  return element as T;
}
