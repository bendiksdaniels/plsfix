// The template catalogue: which ready blocks the pane offers, in the order it
// lists them, and the one lookup the adapter uses. Owns the order and the ids;
// every grid lives in `template-blocks.ts` and the grammar in
// `template-cells.ts`, both of which this module re-exports as one entry point.

import {
  debtSchedule,
  dcf,
  ebitdaBridge,
  npvIrr,
  sensitivity,
  workingCapital,
} from "./template-blocks";
import { type Template } from "./template-cells";

export {
  placeholderRefs,
  resolveFormula,
  type CellRef,
  type Template,
  type TemplateCell,
  type TemplateCellKind,
  type TemplateFormat,
} from "./template-cells";

/** Every template, in the order the pane lists them. */
export const TEMPLATES: Template[] = [
  debtSchedule(),
  dcf(),
  npvIrr(),
  workingCapital(),
  sensitivity(),
  ebitdaBridge(),
];

export function templateById(id: string): Template | null {
  return TEMPLATES.find((template) => template.id === id) ?? null;
}
