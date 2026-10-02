// Workbook project list and the active project new exports inherit. Owns the
// registry fields `projects` and `activeProject`. Invariant: every stored
// name has already been through cleanProject; the two writers run inside
// link-lock's exclusive queue, the same one export/push/remove share.

import type { Registry } from "../link/model";
import {
  cleanProject,
  cleanProjectList,
  withActiveProject,
} from "../link/project";
import { readRegistry, writeRegistry } from "./link-anchors";
import { exclusive } from "./link-lock";

export interface ProjectState {
  names: string[];
  active: string | undefined;
}

export async function readProjectState(): Promise<ProjectState> {
  return Excel.run(async (context) => {
    const registry = await readRegistry(context);
    return {
      names: cleanProjectList([
        ...(registry.projects ?? []),
        ...registry.links
          .map((entry) => entry.project)
          .filter((name): name is string => name !== undefined),
      ]),
      active: registry.activeProject,
    };
  });
}

export async function setActiveProject(
  name: string | undefined,
): Promise<void> {
  await exclusive("project", () =>
    Excel.run(async (context) => {
      const registry = await readRegistry(context);
      writeRegistry(context, asRegistry(withActiveProject(registry, name)));
      await context.sync();
    }),
  );
}

export async function moveLinksToProject(
  ids: readonly string[],
  name: string | undefined,
): Promise<void> {
  const project =
    name === undefined ? undefined : (cleanProject(name) ?? undefined);
  const wanted = new Set(ids);
  await exclusive("project", () =>
    Excel.run(async (context) => {
      const registry = await readRegistry(context);
      const links = registry.links.map((entry) => {
        if (!wanted.has(entry.id)) return entry;
        const next = { ...entry };
        if (project) next.project = project;
        else delete next.project;
        return next;
      });
      writeRegistry(
        context,
        asRegistry(withActiveProject({ ...registry, links }, project)),
      );
      await context.sync();
    }),
  );
}

function asRegistry(value: ReturnType<typeof withActiveProject>): Registry {
  return value as Registry;
}
