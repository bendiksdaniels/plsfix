// Toast: the one status banner shared by every pane. Shows a message for a
// fixed duration, then hides; an error with details renders a "Copy details"
// button so a failure can be pasted into a bug report. No Office.js here -
// this module only ever touches the DOM node it is given.

import { copyText } from "./clipboard";

export type ToastKind = "success" | "error";

export interface Toast {
  show(message: string, kind?: ToastKind, details?: string): void;
}

const DEFAULT_HIDE_AFTER_MS = 3200;

export function createToast(
  container: HTMLElement,
  hideAfterMs = DEFAULT_HIDE_AFTER_MS,
): Toast {
  let timer: number | undefined;

  function show(
    message: string,
    kind: ToastKind = "success",
    details?: string,
  ): void {
    window.clearTimeout(timer);
    container.replaceChildren();

    const text = document.createElement("span");
    text.className = "toast-text";
    text.textContent = message;
    container.append(text);

    if (details) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "toast-copy";
      button.textContent = "Copy details";
      button.addEventListener("click", () => void copyText(details));
      container.append(button);
    }

    container.className = `toast visible ${kind}`;
    timer = window.setTimeout(() => {
      container.className = "toast";
    }, hideAfterMs);
  }

  return { show };
}
