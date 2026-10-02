// The help copy both panes show behind the "?" on a section heading, plus the
// rule for what that copy has to cover. Owns the two keys the whole feature
// turns on: a section's key (its data-help, else its aria-labelledby) and a
// button's key (its data-action, else its element id). The sentences live in
// copy.excel.ts and copy.ppt.ts; the coverage gate is copy.test.ts.

import { EXCEL_HELP } from "./copy.excel";
import { PPT_HELP } from "./copy.ppt";

export interface SectionHelp {
  /** One line on what the section is for. */
  about: string;
  /** One or two plain sentences per button, keyed by data-action or id. */
  buttons: Record<string, string>;
}

/** Section key -> its help. A key means one section across both panes. */
export type HelpCopy = Record<string, SectionHelp>;

/** What one button's sentence may not run past, so the card stays a card. */
export const HELP_MAX_CHARS = 140;

export const HELP: HelpCopy = { ...EXCEL_HELP, ...PPT_HELP };

/** A button the help has to explain, in the order the section shows it. */
export interface HelpButton {
  key: string;
  element: HTMLElement;
}

/** A section the help attaches to: where the toggle goes, and what it covers. */
export interface HelpSection {
  key: string;
  heading: HTMLElement;
  buttons: HelpButton[];
}

// data-action first: a button that carries both (find, share-prepare) is one
// button, and keying it twice would let one of the two keys rot unnoticed.
function buttonKey(button: HTMLElement): string | null {
  return button.dataset.action ?? (button.id === "" ? null : button.id);
}

function buttonsOf(section: Element): HelpButton[] {
  const found: HelpButton[] = [];
  for (const element of section.querySelectorAll<HTMLElement>("button")) {
    const key = buttonKey(element);
    if (key !== null) found.push({ key, element });
  }
  return found;
}

/**
 * Every section of a pane that can carry help, in document order. A section
 * without a heading has nowhere to hang the toggle and is left out.
 */
export function helpSections(root: ParentNode): HelpSection[] {
  const sections: HelpSection[] = [];
  for (const section of root.querySelectorAll("section")) {
    const heading = section.querySelector<HTMLElement>(".section-heading");
    const key =
      section.getAttribute("data-help") ??
      section.getAttribute("aria-labelledby");
    if (heading === null || key === null) continue;
    sections.push({ key, heading, buttons: buttonsOf(section) });
  }
  return sections;
}
