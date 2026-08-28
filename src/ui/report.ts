// Error surface: turns whatever a catch block or a window error event caught
// into a short user-facing message plus a details block worth pasting into a
// bug report. Host-agnostic - it names its host and version through
// ReportContext rather than assuming Excel, so the PowerPoint pane can reuse it.

export interface ReportContext {
  host: string;
  version: string;
}

const FALLBACK_MESSAGE = "The add-in could not complete that action.";

function readProp(error: unknown, key: string): unknown {
  if (typeof error !== "object" || error === null || !(key in error)) return undefined;
  return (error as Record<string, unknown>)[key];
}

export function describeError(
  error: unknown,
  ctx: ReportContext,
  action?: string,
): { message: string; details: string } {
  const message = error instanceof Error ? error.message : FALLBACK_MESSAGE;

  const lines = [String(error)];
  if (action !== undefined) lines.push(`action: ${action}`);
  lines.push(`host: ${ctx.host}`, `version: ${ctx.version}`, `platform: ${navigator.userAgent}`);

  const code = readProp(error, "code");
  if (typeof code === "string") lines.push(`code: ${code}`);

  const debugInfo = readProp(error, "debugInfo");
  if (debugInfo !== undefined) lines.push(`debugInfo: ${JSON.stringify(debugInfo)}`);

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
