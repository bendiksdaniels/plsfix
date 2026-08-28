// Toast: the one status banner shared by every pane. Shows a message for a
// fixed duration, then hides; an error with details renders a "Copy details"
// button so a failure can be pasted into a bug report. No Office.js here -
// this module only ever touches the DOM node it is given.

export type ToastKind = "success" | "error";

export interface Toast {
  show(message: string, kind?: ToastKind, details?: string): void;
}

const DEFAULT_HIDE_AFTER_MS = 3200;

async function copyDetails(details: string): Promise<void> {
  if (navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(details);
      return;
    } catch {
      // Office webviews can deny clipboard access; fall back below.
    }
  }
  copyViaHiddenTextarea(details);
}

// execCommand("copy") only acts on a selected, focused element, so the
// fallback needs a real (if invisible) textarea in the document.
function copyViaHiddenTextarea(details: string): void {
  const area = document.createElement("textarea");
  area.value = details;
  area.style.position = "fixed";
  area.style.opacity = "0";
  document.body.append(area);
  area.select();
  document.execCommand("copy");
  area.remove();
}

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
      button.addEventListener("click", () => void copyDetails(details));
      container.append(button);
    }

    container.className = `toast visible ${kind}`;
    timer = window.setTimeout(() => {
      container.className = "toast";
    }, hideAfterMs);
  }

  return { show };
}
