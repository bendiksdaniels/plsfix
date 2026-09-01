// Everything that formats one shape: its fill, its line, the adjustments a
// geometry exposes and the text frame with the range, font and paragraph under
// it, over the shape the deck holds. Owns nothing; every write lands on the
// shape's own fields. Invariant: each object resolves its shape on every
// access, so a shape deleted under it is seen as gone.

import { normalizeAngle } from "../../src/link/chart-model";
import {
  invalidArgument,
  type FakePptShape,
  type FakePresentation,
} from "./model";
import { FakeClientResult, Loadable } from "./strict";

// What PowerPoint answers when a shape's adjustments are addressed before the
// shape has been synced once: the host has not heard of the id yet.
function unsynced(): Error {
  const error = new Error("InvalidParam passed to GetItem(id)");
  return Object.assign(error, { code: "InvalidParam" });
}

// fill.transparency and lineFormat.weight read back as null once the shape
// stops supporting them (a cleared fill, an invisible line); the d.ts types
// them as a plain number regardless, so a caller that forwards a captured
// null through unchanged reaches this setter, not a compiler error.
function requireFiniteNumber(label: string, value: number): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw invalidArgument(`${label} must be a number`);
  }
  return value;
}

abstract class ShapeFormatBound extends Loadable {
  constructor(
    protected deck: FakePresentation,
    protected shapeId: string,
  ) {
    super();
  }

  protected shape(): FakePptShape {
    return this.deck.findShape(this.shapeId).shape;
  }
}

// The fill/line fields PowerPointApi 1.4 added after fillColor/lineColor were
// modeled: attached directly to the shape object rather than declared on
// model.ts's FakePptShape, which is already at the file's line cap. Reading
// through the same object a copy/paste clone shares - a shallow spread of
// exactly these fields, `cloneShape` in model.ts - carries them across a
// clone the same way it carries every field model.ts does declare.
interface FillLineExtras {
  fillTransparency: number;
  lineTransparency: number;
  lineDashStyle: string;
  lineStyle: string;
}

function extras(shape: FakePptShape): FillLineExtras {
  const extended = shape as FakePptShape & Partial<FillLineExtras>;
  extended.fillTransparency ??= 0;
  extended.lineTransparency ??= 0;
  extended.lineDashStyle ??= "Solid";
  extended.lineStyle ??= "Single";
  return extended as FillLineExtras;
}

// PowerPoint.ShapeFill: the picture a link is painted with, the solid brand
// colour a chart bar takes, and the clear that leaves the shape with neither.
export class ShapeFillProxy extends ShapeFormatBound {
  // Derived, never stored: whichever of fillImage/fillColor is set decides it,
  // the same way real PowerPoint reports a fill type nobody chose directly.
  get type(): string {
    const shape = this.shape();
    if (shape.fillImage !== null) return "PictureAndTexture";
    if (shape.fillColor !== null) return "Solid";
    return "NoFill";
  }
  get foregroundColor(): string {
    return this.shape().fillColor ?? "";
  }
  // PowerPointApi 1.4: null once the fill type no longer supports one, a
  // cleared (NoFill) shape being the one pls,fix ever meets.
  get transparency(): number | null {
    return this.type === "NoFill"
      ? null
      : extras(this.shape()).fillTransparency;
  }
  set transparency(value: number) {
    extras(this.shape()).fillTransparency = requireFiniteNumber(
      "fill.transparency",
      value,
    );
  }
  setImage(base64EncodedImage: string): void {
    const shape = this.shape();
    shape.fillImage = base64EncodedImage;
    shape.setImageCalls += 1;
  }
  setSolidColor(color: string): void {
    const shape = this.shape();
    shape.fillColor = color;
    shape.fillCleared = false;
  }
  clear(): void {
    const shape = this.shape();
    shape.fillColor = null;
    shape.fillImage = null;
    shape.fillCleared = true;
  }
}

export class ShapeLineProxy extends ShapeFormatBound {
  get visible(): boolean {
    return this.shape().lineVisible;
  }
  set visible(value: boolean) {
    this.shape().lineVisible = value;
  }
  get color(): string | null {
    return this.shape().lineColor;
  }
  set color(value: string | null) {
    this.shape().lineColor = value;
  }
  // PowerPointApi 1.4: null while the line isn't visible, whatever weight
  // was last written.
  get weight(): number | null {
    return this.shape().lineVisible ? this.shape().lineWeight : null;
  }
  set weight(value: number) {
    this.shape().lineWeight = requireFiniteNumber("lineFormat.weight", value);
  }
  get transparency(): number {
    return extras(this.shape()).lineTransparency;
  }
  set transparency(value: number) {
    extras(this.shape()).lineTransparency = value;
  }
  // PowerPointApi 1.4: null while the line isn't visible - see weight above.
  get dashStyle(): string | null {
    return this.shape().lineVisible ? extras(this.shape()).lineDashStyle : null;
  }
  set dashStyle(value: string) {
    extras(this.shape()).lineDashStyle = value;
  }
  get style(): string | null {
    return this.shape().lineVisible ? extras(this.shape()).lineStyle : null;
  }
  set style(value: string) {
    extras(this.shape()).lineStyle = value;
  }
}

// PowerPoint.Adjustments (PowerPointApi 1.10): the handles a geometry exposes,
// two on a pie and none on a rectangle. The whole object is unaddressable
// until the shape's first sync, and every value the fake keeps is an angle, so
// a write is normalised the way PowerPoint normalises a pie's.
export class AdjustmentsProxy extends ShapeFormatBound {
  get count(): number {
    return this.values().length;
  }
  get(index: number): FakeClientResult<number> {
    return new FakeClientResult(this.at(index));
  }
  set(index: number, value: number): void {
    this.at(index);
    this.values()[index] = normalizeAngle(value);
  }

  private values(): number[] {
    const shape = this.shape();
    if (!shape.synced) throw unsynced();
    return shape.adjustments;
  }
  private at(index: number): number {
    const value = this.values()[index];
    if (value === undefined) {
      throw invalidArgument(`no adjustment ${String(index)}`);
    }
    return value;
  }
}

// PowerPoint.TextFrame: whether the shape holds any text - which is how an
// empty layout placeholder is told from an object somebody put there - plus
// the box settings a label may write. The insets are left alone by design.
export class TextFrameProxy extends ShapeFormatBound {
  get hasText(): boolean {
    return this.shape().hasText;
  }
  get textRange(): TextRangeProxy {
    return new TextRangeProxy(this.deck, this.shapeId);
  }
  get autoSizeSetting(): string | null {
    return this.shape().autoSize;
  }
  set autoSizeSetting(value: string | null) {
    this.shape().autoSize = value;
  }
  get wordWrap(): boolean | null {
    return this.shape().wordWrap;
  }
  set wordWrap(value: boolean | null) {
    this.shape().wordWrap = value;
  }
  get verticalAlignment(): string | null {
    return this.shape().verticalAlignment;
  }
  set verticalAlignment(value: string | null) {
    this.shape().verticalAlignment = value;
  }
  get leftMargin(): number | null {
    return this.shape().margins.left;
  }
  set leftMargin(value: number | null) {
    this.shape().margins.left = value;
  }
  get rightMargin(): number | null {
    return this.shape().margins.right;
  }
  set rightMargin(value: number | null) {
    this.shape().margins.right = value;
  }
  get topMargin(): number | null {
    return this.shape().margins.top;
  }
  set topMargin(value: number | null) {
    this.shape().margins.top = value;
  }
  get bottomMargin(): number | null {
    return this.shape().margins.bottom;
  }
  set bottomMargin(value: number | null) {
    this.shape().margins.bottom = value;
  }
}

// PowerPoint.TextRange: the fake keeps one run per shape, so the range is the
// whole text and the font and paragraph formats are the shape's.
class TextRangeProxy extends ShapeFormatBound {
  get text(): string {
    return this.shape().text ?? "";
  }
  set text(value: string) {
    const shape = this.shape();
    shape.text = value;
    shape.hasText = value.length > 0;
  }
  get font(): TextFontProxy {
    return new TextFontProxy(this.deck, this.shapeId);
  }
  get paragraphFormat(): ParagraphFormatProxy {
    return new ParagraphFormatProxy(this.deck, this.shapeId);
  }
}

class TextFontProxy extends ShapeFormatBound {
  get name(): string | undefined {
    return this.shape().font.name;
  }
  set name(value: string | undefined) {
    this.shape().font.name = value;
  }
  get size(): number | undefined {
    return this.shape().font.size;
  }
  set size(value: number | undefined) {
    this.shape().font.size = value;
  }
  get color(): string | undefined {
    return this.shape().font.color;
  }
  set color(value: string | undefined) {
    this.shape().font.color = value;
  }
  get bold(): boolean | undefined {
    return this.shape().font.bold;
  }
  set bold(value: boolean | undefined) {
    this.shape().font.bold = value;
  }
}

class ParagraphFormatProxy extends ShapeFormatBound {
  get horizontalAlignment(): string | null {
    return this.shape().alignment;
  }
  set horizontalAlignment(value: string | null) {
    this.shape().alignment = value;
  }
}
