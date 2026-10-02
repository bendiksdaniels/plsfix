// src/ui/clipboard-links.ts
// The clipboard for link bundles: text/html carries the bundle in one data
// attribute, text/plain only the sentence, so a paste anywhere but the
// pls,fix pane gives one line. Invariant: the write starts inside the press
// (WebKit requires it); the content may arrive after the Excel work.

import {
  BUNDLE_SENTENCE,
  bundleHtml,
  readPastedBundle,
  type BundleRead,
} from "../link/bundle";

export interface PendingCopy {
  resolve(json: string): void;
  reject(error: unknown): void;
  // True once the write landed; false when the engine has no ClipboardItem
  // or refused it (focus moved, a permission): the caller then offers the
  // synchronous copy behind a second press.
  done: Promise<boolean>;
}

export function beginClipboardWrite(): PendingCopy {
  let resolveJson!: (json: string) => void;
  let rejectJson!: (error: unknown) => void;
  const json = new Promise<string>((resolve, reject) => {
    resolveJson = resolve;
    rejectJson = reject;
  });
  json.catch(() => undefined);
  const Item = (globalThis as { ClipboardItem?: typeof ClipboardItem })
    .ClipboardItem;
  const clipboard = (globalThis as { navigator?: Navigator }).navigator
    ?.clipboard;
  let done: Promise<boolean> = Promise.resolve(false);
  if (Item !== undefined && clipboard?.write !== undefined) {
    try {
      const html = json.then(
        (text) => new Blob([bundleHtml(text)], { type: "text/html" }),
      );
      const plain = json.then(
        () => new Blob([BUNDLE_SENTENCE], { type: "text/plain" }),
      );
      done = clipboard
        .write([new Item({ "text/html": html, "text/plain": plain })])
        .then(
          () => true,
          () => false,
        );
    } catch {
      done = Promise.resolve(false);
    }
  }
  return { resolve: resolveJson, reject: rejectJson, done };
}

// Inside a press only: execCommand("copy") is refused anywhere else.
export function copyBundleNow(json: string): boolean {
  let wrote = false;
  const onCopy = (event: Event): void => {
    const data = (event as ClipboardEvent).clipboardData;
    if (!data) return;
    data.setData("text/html", bundleHtml(json));
    data.setData("text/plain", BUNDLE_SENTENCE);
    event.preventDefault();
    wrote = true;
  };
  document.addEventListener("copy", onCopy);
  try {
    return document.execCommand("copy") && wrote;
  } catch {
    return false;
  } finally {
    document.removeEventListener("copy", onCopy);
  }
}

// Read inside the paste event: its DataTransfer is empty once the event ends.
export function bundleFromPaste(data: DataTransfer | null): BundleRead {
  if (data === null) return { ok: false, reason: "notBundle" };
  return readPastedBundle(
    data.getData("text/html"),
    data.getData("text/plain"),
  );
}
