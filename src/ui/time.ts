// Relative-time formatting shared by the PowerPoint pane (link ages, inbox
// timestamps) and the Excel Links tab (push ages). A null timestamp reads as
// "never"; a clock ahead of the server reads "just now" rather than a
// negative age.

const MINUTE = 60;
const HOUR = 60 * MINUTE;
const DAY = 24 * HOUR;

export const NEVER = "never";

export function relativeTime(
  unixSeconds: number | null,
  now: number = Date.now() / 1000,
): string {
  if (unixSeconds === null) return NEVER;
  const age = Math.max(0, Math.round(now - unixSeconds));
  if (age < MINUTE) return "just now";
  if (age < HOUR) return `${String(Math.floor(age / MINUTE))} min ago`;
  if (age < DAY) return `${String(Math.floor(age / HOUR))} h ago`;
  return counted(Math.floor(age / DAY), "day");
}

function counted(count: number, unit: string): string {
  return `${String(count)} ${unit}${count === 1 ? "" : "s"} ago`;
}

// Both link lists carry an ISO stamp - the registry's last push, the inbox
// item's creation - and want the same age line, so the parse lives here rather
// than being repeated in each renderer. A stamp that will not parse reads as
// no stamp at all.
export function relativeStamp(iso: string | null): string {
  if (iso === null) return NEVER;
  const seconds = Date.parse(iso) / 1000;
  return Number.isFinite(seconds) ? relativeTime(seconds) : NEVER;
}
