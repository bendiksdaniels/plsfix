// Two-click confirm for a destructive or one-way button: the first press
// only relabels it and arms a CONFIRM_MS window, a second press inside that
// window runs the action, and any later press or the window lapsing on its
// own puts the label back. One helper owns the timer, the "armed" class, the
// relabel and the second press, so every confirm button in both panes -
// styles-delete and delete-names first, the rest of slice W after them -
// behaves alike.
//
// The relabel targets the button's own <strong> when it has one (the rich
// action-list rows, e.g. clean-past-data): the icon and the description
// beside it survive untouched. An icon-only button (class "icon-button",
// e.g. reset-brand) keeps its glyph and relabels through its title and
// aria-label instead, so a fixed-size button never stretches to fit a
// sentence. Every other button gets its whole content swapped, the way
// styles-delete and delete-names already did.

export const CONFIRM_MS = 5_000;

// No trailing full stop: this is a short button label, not a sentence (the
// two buttons this pattern started on never had one either).
const CONFIRM_LABEL = "Click again to confirm";

export interface ConfirmButton {
  isArmed(): boolean;
  arm(): void;
  disarm(): void;
  // Wire directly as the button's click listener: arms on the first press,
  // disarms and runs on the second, or runs at once while `when` says this
  // press needs no confirming.
  handleClick(): void;
}

export interface ArmConfirmOptions {
  // Answers whether THIS press needs confirming at all. False skips the arm
  // window and runs immediately (generate-key: a first key needs no
  // confirming, only replacing one does).
  when?: () => boolean;
  // The resting label once nothing is armed, read fresh every time rather
  // than frozen at construction - styles-delete and delete-names show a
  // live count that a rescan between presses can change.
  label?: () => string;
}

export function armConfirm(
  button: HTMLButtonElement,
  run: () => void,
  options: ArmConfirmOptions = {},
): ConfirmButton {
  const iconOnly = button.classList.contains("icon-button");
  const target: HTMLElement = button.querySelector("strong") ?? button;
  const restingHTML = button.innerHTML;
  const restingAriaLabel = button.getAttribute("aria-label");
  const restingTitle = button.getAttribute("title");

  let armed = false;
  let timer: number | undefined;

  function paint(armedNow: boolean): void {
    if (iconOnly) {
      // The glyph itself never changes size or shape; only what a screen
      // reader and a hover tooltip say does, plus the "armed" class below.
      if (restingAriaLabel !== null) {
        button.setAttribute(
          "aria-label",
          armedNow ? CONFIRM_LABEL : restingAriaLabel,
        );
      }
      if (restingTitle !== null) {
        button.setAttribute("title", armedNow ? CONFIRM_LABEL : restingTitle);
      }
      return;
    }
    if (armedNow) {
      target.textContent = CONFIRM_LABEL;
      return;
    }
    if (options.label) target.textContent = options.label();
    else button.innerHTML = restingHTML;
  }

  function disarm(): void {
    window.clearTimeout(timer);
    armed = false;
    button.classList.remove("armed");
    paint(false);
  }

  function arm(): void {
    armed = true;
    button.classList.add("armed");
    paint(true);
    timer = window.setTimeout(disarm, CONFIRM_MS);
  }

  function handleClick(): void {
    if (options.when && !options.when()) {
      run();
      return;
    }
    if (!armed) {
      arm();
      return;
    }
    disarm();
    run();
  }

  return { isArmed: () => armed, arm, disarm, handleClick };
}
