// The one place a negative shape size is refused: PowerPoint throws
// InvalidArgument for a negative width or height, both on ShapeAddOptions
// and on the Shape width/height setters, so objects.ts (already at the file
// line cap) routes both through this instead of growing past it.

import { invalidArgument, type FakeShapeInit } from "./model";

// Shape.width / Shape.height: hands the value back so a setter can assign
// through the call in one expression.
export function requireNonNegative(value: number): number {
  if (value < 0) throw invalidArgument("size must not be negative");
  return value;
}

// ShapeAddOptions: width/height are optional at add time, so only a side the
// caller actually gave is checked.
export function validatedShapeInit(init: FakeShapeInit): FakeShapeInit {
  if (init.width !== undefined) requireNonNegative(init.width);
  if (init.height !== undefined) requireNonNegative(init.height);
  return init;
}
