// The PowerPoint pane's help copy, in the same shape as the Excel half. Owns
// the wording of the deck side only. Its Linked objects section is keyed
// "deck-links" through a data-help attribute, because the Excel pane already
// owns "links-heading" and one key may mean one section.

import type { HelpCopy } from "./copy";

export const PPT_HELP: HelpCopy = {
  "deck-links": {
    about:
      "Every pls,fix object in this deck, and the buttons that refresh or release them.",
    buttons: {
      "refresh-links":
        "Scans the deck again and asks the relay which links have a newer picture waiting.",
      "update-selected":
        "Repaints the ticked links with their latest picture. Slide and position do not move.",
      "update-slide": "Repaints every link on the slide you are on.",
      "update-all": "Repaints every link in this deck.",
      "revert-selected":
        "Puts the ticked links back to the picture before the last update. One step only.",
      "go-to-slide": "Jumps to the slide the first ticked link sits on.",
      "change-source":
        "Points one ticked link at another export waiting in the Inbox. Tick exactly one row.",
      "change-source-confirm":
        "Re-points the link at the chosen export. Slide, position and size stay as they are.",
      "change-source-cancel":
        "Closes the picker and leaves the link pointing where it did.",
      "break-selected":
        "Takes the link marker off the ticked objects. They stay on the slide but stop updating.",
    },
  },
  "inbox-heading": {
    about:
      "Exports from Excel not yet in this deck; they expire after 7 days. A chart lands as editable shapes, else a picture.",
    buttons: {
      "refresh-inbox":
        "Asks the relay for the exports waiting under your link key.",
      "first-run-ppt-dismiss":
        "Dismisses this card. It will not show again on this machine.",
    },
  },
  "pairing-heading": {
    about:
      "Paste the key from Excel once per computer. It stays here, and the relay never sees it.",
    buttons: {
      "save-key":
        "Stores the pasted key on this computer and opens the inbox items pushed from Excel.",
      "forget-key":
        "Deletes the key from this computer. The deck's links stay, but the inbox closes.",
    },
  },
};
