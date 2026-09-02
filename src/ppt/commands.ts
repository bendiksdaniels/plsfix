// Ribbon commands of the PowerPoint pane (shared runtime): one table from
// every PLSFIX_PPT_* FunctionName in manifest/spec.ts to the object tool it
// runs, with its own promise chain per command because a ribbon button fires
// with the pane closed too. Office.js only reaches here through ./object-tools.

import { describeError } from "../ui/report";
import type { ToastKind } from "../ui/toast";
import type { AlignMode } from "./object-math";
import {
  alignSelected,
  applyObjectStyle,
  captureObjectStyle,
  distributeSelected,
  matchSelectedSize,
  selectSimilar,
  swapSelected,
} from "./object-tools";

interface CommandEvent {
  completed: () => void;
}

export interface CommandDeps {
  // The pane's toast; a command's one-line result lands there like a click's.
  notify: (message: string, kind?: ToastKind, details?: string) => void;
  context: { host: string; version: string };
  // The ribbon's "Object tools" button: the pane up, on its Tools tab.
  showTools: () => Promise<void>;
}

const ALIGN_COMMANDS: readonly (readonly [string, AlignMode])[] = [
  ["PLSFIX_PPT_ALIGN_LEFT", "left"],
  ["PLSFIX_PPT_ALIGN_CENTER", "center"],
  ["PLSFIX_PPT_ALIGN_RIGHT", "right"],
  ["PLSFIX_PPT_ALIGN_TOP", "top"],
  ["PLSFIX_PPT_ALIGN_MIDDLE", "middle"],
  ["PLSFIX_PPT_ALIGN_BOTTOM", "bottom"],
];

// Every FunctionName the Presentation host declares, and nothing else:
// manifest/xml.test.ts holds the two lists to each other.
export function commandTable(
  deps: CommandDeps,
): Record<string, () => Promise<string>> {
  const table: Record<string, () => Promise<string>> = {
    PLSFIX_PPT_TOOLS: async () => {
      await deps.showTools();
      return "";
    },
    PLSFIX_PPT_MATCH: matchSelectedSize,
    PLSFIX_PPT_SIMILAR: selectSimilar,
    PLSFIX_PPT_SWAP: swapSelected,
    PLSFIX_PPT_CAPTURE: captureObjectStyle,
    PLSFIX_PPT_PAINT: applyObjectStyle,
    PLSFIX_PPT_DIST_ACROSS: () => distributeSelected("horizontal"),
    PLSFIX_PPT_DIST_DOWN: () => distributeSelected("vertical"),
  };
  for (const [id, mode] of ALIGN_COMMANDS) {
    table[id] = () => alignSelected(mode);
  }
  return table;
}

export function registerCommands(deps: CommandDeps): void {
  if (!Office.actions?.associate) return;
  for (const [id, run] of Object.entries(commandTable(deps))) {
    Office.actions.associate(id, (event?: CommandEvent) => {
      void run()
        .then((message) => {
          if (message !== "") deps.notify(message, "success");
        })
        .catch((error: unknown) => {
          const { message, details } = describeError(error, deps.context, id);
          deps.notify(message, "error", details);
        })
        .finally(() => event?.completed());
    });
  }
}
