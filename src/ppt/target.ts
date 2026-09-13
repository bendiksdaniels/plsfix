// Reads the inbox's Slide and Where pickers into an InsertTarget, and keeps
// the Slide picker's options in step with the deck's slide count. Kept out
// of main.ts, already over its size cap: that file only ever calls in here.

import { slideCount, slideIdAt } from "./host";
import type { InsertTarget, Where } from "./placement";
import { renderSlideOptions } from "./views";

// Re-synced with the inbox, and once at boot: on a read the host refuses,
// the picker keeps (or falls back to) "This slide" alone rather than options
// that may no longer match the deck.
export async function refreshSlideOptions(
  select: HTMLSelectElement,
): Promise<void> {
  try {
    renderSlideOptions(select, await slideCount());
  } catch {
    renderSlideOptions(select, 0);
  }
}

// "" (This slide) is slideId null, same as every insert before this picker
// existed; any other value names a position, resolved to that slide's actual
// id right here so what an insert reads is never a stale one held since the
// select was last rendered.
export async function readInsertTarget(
  slideSelect: HTMLSelectElement,
  whereSelect: HTMLSelectElement,
): Promise<InsertTarget> {
  const picked = slideSelect.value;
  const slideId = picked === "" ? null : await slideIdAt(Number(picked) - 1);
  return { slideId, where: whereSelect.value as Where };
}
