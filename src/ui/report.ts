// Error surface: turns whatever a catch block or a window error event caught
// into a short user-facing message plus a details block worth pasting into a
// bug report. Host-agnostic - it names its host and version through
// ReportContext rather than assuming Excel, so the PowerPoint pane can reuse it.
// A read sync in src/excel is a bare host string by design (only a write
// stages its own refusal, through syncWrite/paintSync), so a HOST error - one
// carrying an office.js `code` - whose message names no stage is prefixed with
// the running action's label. The pane's own sentences are left alone.

export interface ReportContext {
  host: string;
  version: string;
}

const FALLBACK_MESSAGE = "The add-in could not complete that action.";

function readProp(error: unknown, key: string): unknown {
  if (typeof error !== "object" || error === null || !(key in error))
    return undefined;
  return (error as Record<string, unknown>)[key];
}

/** The action id in words: hyphens to spaces, first letter capitalised. */
function actionLabel(action: string): string {
  const words = action.replaceAll("-", " ");
  return words.charAt(0).toUpperCase() + words.slice(1);
}

// A message already names its own stage - "Indent cycling: this sheet is
// protected" (a colon) or "Pinstripes need at least two rows..." (opens with
// the label's own first word) - so prefixing it again would repeat itself.
function hasOwnStage(message: string, label: string): boolean {
  if (message.includes(": ")) return true;
  const firstWord = label.split(" ")[0] ?? label;
  return message.toLowerCase().startsWith(firstWord.toLowerCase());
}

function withActionLabel(message: string, action: string): string {
  const label = actionLabel(action);
  return hasOwnStage(message, label) ? message : `${label}: ${message}`;
}

// Office.js errors carry a string `code` (GeneralException, AccessDenied...);
// an Error the pane threw itself never does.
function isHostError(error: unknown): boolean {
  return typeof readProp(error, "code") === "string";
}

export function describeError(
  error: unknown,
  ctx: ReportContext,
  action?: string,
): { message: string; details: string } {
  const raw = error instanceof Error ? error.message : FALLBACK_MESSAGE;
  const message =
    action !== undefined && isHostError(error)
      ? withActionLabel(raw, action)
      : raw;

  const lines = [String(error)];
  if (action !== undefined) lines.push(`action: ${action}`);
  lines.push(
    `host: ${ctx.host}`,
    `version: ${ctx.version}`,
    `platform: ${navigator.userAgent}`,
  );

  const code = readProp(error, "code");
  if (typeof code === "string") lines.push(`code: ${code}`);

  const debugInfo = readProp(error, "debugInfo");
  if (debugInfo !== undefined)
    lines.push(`debugInfo: ${JSON.stringify(debugInfo)}`);

  if (error instanceof Error && error.stack) lines.push(error.stack);

  return { message, details: lines.join("\n") };
}

export function installErrorReporting(
  ctx: ReportContext,
  notify: (message: string, details: string) => void,
  win: Window = window,
): void {
  win.addEventListener("error", (event) => {
    const error: unknown = event.error;
    const { message, details } = describeError(error, ctx);
    notify(message, details);
  });

  win.addEventListener("unhandledrejection", (event) => {
    const error: unknown = event.reason;
    const { message, details } = describeError(error, ctx);
    notify(message, details);
  });
}
