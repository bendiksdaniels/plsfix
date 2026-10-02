// The "?" on every section heading and the card it opens: what the section is
// for, then one line per button. Owns the toggle, the card and the rule that
// one card is open at a time; the sentences are data in src/help/copy.ts and
// the labels are read off the buttons themselves, so they can never drift.
// Nothing is remembered - closing the pane closes the help with it.

import {
  HELP,
  helpSections,
  type HelpButton,
  type HelpSection,
  type SectionHelp,
} from "../help/copy";

interface OpenCard {
  toggle: HTMLElement;
  card: HTMLElement;
}

// What the button reads as: its aria-label when the face is a glyph, else the
// bold line of an action row, else whatever text it carries.
function labelOf(button: HTMLElement): string {
  const face =
    button.getAttribute("aria-label") ??
    button.querySelector("strong")?.textContent ??
    button.textContent ??
    "";
  return face.replace(/\s+/g, " ").trim() || button.id;
}

function titleOf(section: HelpSection): string {
  const heading = section.heading.querySelector("h2")?.textContent ?? "";
  return heading.replace(/\s+/g, " ").trim() || section.key;
}

function entry(button: HelpButton, sentence: string): HTMLElement {
  const row = document.createElement("div");
  row.className = "help-row";
  const term = document.createElement("dt");
  term.textContent = labelOf(button.element);
  const detail = document.createElement("dd");
  detail.textContent = sentence;
  row.append(term, detail);
  return row;
}

function buildCard(
  id: string,
  section: HelpSection,
  copy: SectionHelp,
): HTMLElement {
  const card = document.createElement("div");
  card.id = id;
  card.className = "help-card";
  card.hidden = true;
  const about = document.createElement("p");
  about.className = "help-about";
  about.textContent = copy.about;
  card.append(about);

  const list = document.createElement("dl");
  list.className = "help-list";
  for (const button of section.buttons) {
    const sentence = copy.buttons[button.key];
    if (sentence !== undefined) list.append(entry(button, sentence));
  }
  if (list.childElementCount > 0) card.append(list);
  return card;
}

function buildToggle(id: string, cardId: string, title: string): HTMLElement {
  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.id = id;
  toggle.className = "help-toggle";
  toggle.setAttribute("aria-expanded", "false");
  toggle.setAttribute("aria-controls", cardId);
  toggle.setAttribute("aria-label", `About ${title}`);
  const mark = document.createElement("span");
  mark.textContent = "?";
  toggle.append(mark);
  return toggle;
}

/**
 * Adds the "?" toggle and its card to every section of `root` the copy covers.
 * A section with no entry is left untouched - the copy gate is what reports it,
 * not the pane at runtime.
 */
export function installHelp(root: ParentNode): void {
  let open: OpenCard | null = null;

  function close(): void {
    if (open === null) return;
    open.card.hidden = true;
    open.toggle.setAttribute("aria-expanded", "false");
    open = null;
  }

  for (const section of helpSections(root)) {
    const copy = HELP[section.key];
    if (copy === undefined) continue;
    const cardId = `help-${section.key}`;
    const card = buildCard(cardId, section, copy);
    const toggle = buildToggle(`${cardId}-toggle`, cardId, titleOf(section));
    toggle.addEventListener("click", () => {
      const wasOpen = open?.toggle === toggle;
      close();
      if (wasOpen) return;
      card.hidden = false;
      toggle.setAttribute("aria-expanded", "true");
      open = { toggle, card };
    });
    section.heading.append(toggle);
    section.heading.insertAdjacentElement("afterend", card);
  }

  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape" || open === null) return;
    const toggle = open.toggle;
    close();
    toggle.focus();
  });
}
