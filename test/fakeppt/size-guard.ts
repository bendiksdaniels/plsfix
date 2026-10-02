// The one place a negative shape size is refused: PowerPoint throws
// InvalidArgument for a negative width or height, both on ShapeAddOptions
// and on the Shape width/height setters, so objects.ts (already at the file
// line cap) routes both through this instead of growing past it. Also the
// addTable column-width check below, for the same file-cap reason.

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

// A thousandth of a point tells a real mismatch from float noise (two sides
// computed the same fractional total two different ways still agree to far
// tighter than this).
const COLUMN_WIDTH_TOLERANCE = 0.001;

// PowerPoint for the web, rig 27.09: ShapeCollection.addTable answers
// InvalidArgument whenever the columns' columnWidth values do not add up
// EXACTLY to the table's own width - proven on the real host by probing a
// table whose columns were rounded to whole points while its own width
// stayed fractional. A column left with no columnWidth (null, "left to
// PowerPoint") takes this out of scope: nothing here can say what that
// column will measure, so the sum is not checked at all.
export function requireColumnWidthsMatch(
  width: number | undefined,
  columns: { columnWidth?: number }[] | undefined,
): void {
  if (width === undefined || columns === undefined) return;
  const widths = columns.map((column) => column.columnWidth);
  if (!widths.every((one): one is number => typeof one === "number")) return;
  const sum = widths.reduce((total, one) => total + one, 0);
  if (Math.abs(sum - width) > COLUMN_WIDTH_TOLERANCE) {
    throw invalidArgument(
      "the sum of the columns' widths must equal the table's own width",
    );
  }
}
