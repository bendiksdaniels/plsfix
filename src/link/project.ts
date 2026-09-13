// The project name a link carries: chosen in Excel, stamped on the registry
// entry, the inbox item and the PowerPoint tag. The relay never sees it.
// Invariant: a missing or empty name is "No project"; a stored name is already
// clean (trimmed, no control characters, at most PROJECT_MAX).

export const PROJECT_MAX = 40;
export const NO_PROJECT = "No project";

const CONTROLS = /[\u0000-\u001f\u007f]/g;

export function cleanProject(name: string): string | null {
  const trimmed = name.replace(CONTROLS, "").trim();
  if (trimmed.length === 0) return null;
  return trimmed.slice(0, PROJECT_MAX);
}

export function projectLabel(name: string | undefined): string {
  return name ?? NO_PROJECT;
}

// A field that is absent is fine; a present one must be a string. Cleaning
// happens in the decoder so a dirty name never drops the whole record.
export function optionalProject(value: unknown): boolean {
  return value === undefined || typeof value === "string";
}

export function readProject(value: unknown): string | undefined {
  return typeof value === "string"
    ? (cleanProject(value) ?? undefined)
    : undefined;
}

export function withProject<T extends { project?: string }>(record: T): T {
  const project = readProject(record.project);
  if (project === undefined) {
    const copy = { ...record };
    delete copy.project;
    return copy;
  }
  return { ...record, project };
}

export function cleanProjectList(
  names: readonly string[] | undefined,
): string[] {
  if (names === undefined) return [];
  const out: string[] = [];
  for (const name of names) {
    const cleaned = readProject(name);
    if (cleaned !== undefined && !out.includes(cleaned)) out.push(cleaned);
  }
  return out;
}

export function withActiveProject(
  registry: {
    v: 1;
    links: { project?: string }[];
    projects?: string[];
    activeProject?: string;
  },
  name: string | undefined,
): typeof registry {
  const activeProject =
    name === undefined ? undefined : (cleanProject(name) ?? undefined);
  const fromLinks = registry.links
    .map((entry) => entry.project)
    .filter((project): project is string => project !== undefined);
  const projects = cleanProjectList([
    ...(registry.projects ?? []),
    ...fromLinks,
    ...(activeProject ? [activeProject] : []),
  ]);
  const next = { ...registry, links: registry.links };
  if (projects.length > 0) next.projects = projects;
  else delete next.projects;
  if (activeProject) next.activeProject = activeProject;
  else delete next.activeProject;
  return next;
}
