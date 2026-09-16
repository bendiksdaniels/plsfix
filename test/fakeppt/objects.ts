// The office.js objects the fake PowerPoint host hands out: slides, the shape
// collections and the adds that fill them, shapes, groups and tags, over the
// deck in model.ts (the formats hanging off a shape live in format.ts). Owns
// nothing itself. Invariant: every object addresses its item by id and reads
// the deck on each access, so it sees a move or a delete happen under it.

import {
  AdjustmentsProxy,
  ShapeFillProxy,
  ShapeLineProxy,
  TextFrameProxy,
} from "./format";
import {
  gone,
  invalidArgument,
  type FakePptShape,
  type FakePresentation,
  type FakeShapeGroup,
  type FakeShapeInit,
  type FakeSlide,
  type ShapeSite,
} from "./model";
import { requireNonNegative, validatedShapeInit } from "./size-guard";
import { FakeClientResult, Loadable } from "./strict";
import { applyPendingTableStyle, newFakeTable, TableProxy } from "./tables";

// The side PowerPoint gives a line whose width or height was left at zero.
const LINE_DEFAULT_SIDE = 72;

// PowerPoint reads a shape group through the group shape, so anything else
// answers the same GeneralException the real host does.
function notAGroup(shapeId: string): Error {
  const error = new Error(`GeneralException: shape "${shapeId}" is no group`);
  return Object.assign(error, { code: "GeneralException" });
}

// getItem addresses one collection: a shape inside a group is not on the
// slide, and a slide's collection says ItemNotFound for it.
function member(shapes: FakePptShape[], id: string): string {
  if (!shapes.some((shape) => shape.id === id)) throw gone("shape", id);
  return id;
}

function groupOf(deck: FakePresentation, shapeId: string): FakeShapeGroup {
  const group = deck.findShape(shapeId).shape.group;
  if (!group) throw notAGroup(shapeId);
  return group;
}

interface BoxOptions {
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

// getSelectedShapes(), in order; off FakePresentation itself, at its cap.
const selectedShapes = new WeakMap<FakePresentation, string[]>();

export function selectedShapeIds(deck: FakePresentation): string[] {
  return selectedShapes.get(deck) ?? [];
}

export function setSelectedShapeIds(
  deck: FakePresentation,
  ids: string[],
): void {
  selectedShapes.set(deck, [...ids]);
}

// Slides, shapes and tags are all addressed by id: the deck holds the truth,
// the id is the handle, and nullable says whether a missing item reports
// isNullObject or throws the way getItem does.
abstract class Handle extends Loadable {
  constructor(
    protected deck: FakePresentation,
    protected handleId: string,
    protected nullable = false,
  ) {
    super();
  }
}

export class SlideCollectionProxy extends Loadable {
  constructor(
    private deck: FakePresentation,
    private list: () => FakeSlide[],
  ) {
    super();
  }

  get items(): SlideProxy[] {
    return this.list().map((slide) => new SlideProxy(this.deck, slide.id));
  }
  getItem(id: string): SlideProxy {
    return new SlideProxy(this.deck, id);
  }
  // By position in the deck, resolved to the slide's id there and then, which
  // is what a later delete or reorder makes the object see.
  getItemAt(index: number): SlideProxy {
    return new SlideProxy(this.deck, this.list()[index]?.id ?? "");
  }
  getItemOrNullObject(id: string): SlideProxy {
    return new SlideProxy(this.deck, id, true);
  }
  getCount(): FakeClientResult<number> {
    return new FakeClientResult(this.list().length);
  }
}

class SlideProxy extends Handle {
  get id(): string {
    return this.nullable && !this.peek() ? "" : this.model().id;
  }
  get index(): number {
    return this.deck.slides.indexOf(this.model());
  }
  get isNullObject(): boolean {
    return this.peek() === null;
  }
  get shapes(): ShapeCollectionProxy {
    return new ShapeCollectionProxy(this.deck, this.handleId);
  }

  // setSelectedShapes: every id must be a shape on this slide, or the whole
  // call is refused; otherwise it replaces the selection and selects the slide.
  setSelectedShapes(shapeIds: string[]): void {
    const onSlide = new Set(this.model().shapes.map((shape) => shape.id));
    for (const id of shapeIds) {
      if (!onSlide.has(id)) throw invalidArgument(`shape "${id}" not here`);
    }
    this.deck.selectedSlideIds = [this.handleId];
    setSelectedShapeIds(this.deck, shapeIds);
  }

  private peek(): FakeSlide | null {
    return this.deck.findSlide(this.handleId);
  }
  private model(): FakeSlide {
    return this.deck.findSlideOrThrow(this.handleId);
  }
}

class ShapeCollectionProxy extends Handle {
  get items(): ShapeProxy[] {
    const shapes = this.slide().shapes;
    return shapes.map((shape) => new ShapeProxy(this.deck, shape.id));
  }
  getItem(id: string): ShapeProxy {
    return new ShapeProxy(this.deck, member(this.slide().shapes, id));
  }
  getItemOrNullObject(id: string): ShapeProxy {
    return new ShapeProxy(this.deck, id, true);
  }
  getCount(): FakeClientResult<number> {
    return new FakeClientResult(this.slide().shapes.length);
  }

  addGeometricShape(geometry: string, options: BoxOptions = {}): ShapeProxy {
    return this.add(geometry, { type: "GeometricShape", geometry, ...options });
  }

  // PowerPoint.ShapeCollection.addTextBox: the box arrives with its text in
  // it, sized by the caller, with PowerPoint's own insets and font.
  addTextBox(text: string, options: BoxOptions = {}): ShapeProxy {
    return this.add("TextBox", { type: "TextBox", text, ...options });
  }

  // addLine: the connector type is the line's geometry, the box its ends. The
  // host reads a zero width or height as "not given" and uses its default,
  // which is how a baseline asked for at height 0 comes out sloped (measured
  // on PowerPoint for the web, 30.08); a later `shape.height = 0` sticks.
  addLine(connectorType = "Straight", options: BoxOptions = {}): ShapeProxy {
    return this.add("Line", {
      type: "Line",
      geometry: connectorType,
      ...options,
      width: options.width || LINE_DEFAULT_SIDE,
      height: options.height || LINE_DEFAULT_SIDE,
    });
  }

  // addGroup: the members leave the slide's top level for the new group, whose
  // box is the one they occupy together. Fewer than two ids is refused.
  addGroup(shapeIds: string[]): ShapeProxy {
    const group = this.deck.groupShapes(shapeIds, this.slide().id);
    return new ShapeProxy(this.deck, group.id);
  }

  // PowerPoint.ShapeCollection.addTable (PowerPointApi 1.8): the grid arrives
  // with its text already in it, and every format inherited from the style.
  addTable(
    rowCount: number,
    columnCount: number,
    options: BoxOptions & {
      values?: string[][];
      columns?: { columnWidth?: number }[];
    } = {},
  ): ShapeProxy {
    const slide = this.slide();
    const { values, columns, ...box } = options;
    const shape = this.deck.addShape(slide, {
      name: `Table ${String(slide.shapes.length + 1)}`,
      type: "Table",
      ...box,
    });
    shape.table = newFakeTable(rowCount, columnCount, values);
    shape.table.columnWidths = (columns ?? []).map(
      (column) => column.columnWidth ?? null,
    );
    applyPendingTableStyle(this.deck, shape);
    return new ShapeProxy(this.deck, shape.id);
  }

  // PowerPoint names a new shape after its kind and its place in the stack,
  // "Rectangle 3", whatever the add was.
  private add(label: string, init: FakeShapeInit): ShapeProxy {
    const slide = this.slide();
    const shape = this.deck.addShape(slide, {
      name: `${label} ${String(slide.shapes.length + 1)}`,
      ...validatedShapeInit(init),
    });
    return new ShapeProxy(this.deck, shape.id);
  }

  private slide(): FakeSlide {
    return this.deck.findSlideOrThrow(this.handleId);
  }
}

// The shape itself, its fill, its line format and its tags all hang off one
// entry in the deck, and every one of them resolves it again on each access.
abstract class ShapeBound extends Handle {
  protected peek(): ShapeSite | null {
    return this.deck.peekShape(this.handleId);
  }
  protected model(): FakePptShape {
    return this.deck.findShape(this.handleId).shape;
  }
}

export class ShapeProxy extends ShapeBound {
  // The id is the object's own handle, so it survives the shape being deleted.
  get id(): string {
    return this.nullable && !this.peek() ? "" : this.handleId;
  }
  get isNullObject(): boolean {
    return this.peek() === null;
  }
  get type(): string {
    return this.model().type;
  }
  get zOrderPosition(): number {
    return this.model().zOrder;
  }
  get name(): string {
    return this.model().name;
  }
  set name(value: string) {
    this.model().name = value;
  }
  get left(): number {
    return this.model().left;
  }
  set left(value: number) {
    this.model().left = value;
  }
  get top(): number {
    return this.model().top;
  }
  set top(value: number) {
    this.model().top = value;
  }
  get width(): number {
    return this.model().width;
  }
  set width(value: number) {
    this.model().width = requireNonNegative(value);
  }
  get height(): number {
    return this.model().height;
  }
  set height(value: number) {
    this.model().height = requireNonNegative(value);
  }

  get fill(): ShapeFillProxy {
    return new ShapeFillProxy(this.deck, this.handleId);
  }
  get lineFormat(): ShapeLineProxy {
    return new ShapeLineProxy(this.deck, this.handleId);
  }
  get textFrame(): TextFrameProxy {
    return new TextFrameProxy(this.deck, this.handleId);
  }
  get adjustments(): AdjustmentsProxy {
    return new AdjustmentsProxy(this.deck, this.handleId);
  }
  get tags(): TagCollectionProxy {
    return new TagCollectionProxy(this.deck, this.handleId);
  }
  get group(): ShapeGroupProxy {
    return new ShapeGroupProxy(this.deck, this.handleId);
  }

  getTable(): TableProxy {
    return new TableProxy(this.deck, this.handleId);
  }
  delete(): void {
    this.deck.deleteShape(this.handleId);
  }
  getParentSlideOrNullObject(): SlideProxy {
    return new SlideProxy(this.deck, this.peek()?.slide.id ?? "", true);
  }
  // getParentSlide: the slide proxy, or the ItemNotFound a gone shape answers
  // everywhere else in this fake.
  getParentSlide(): SlideProxy {
    const site = this.peek();
    if (!site) throw gone("shape", this.handleId);
    return new SlideProxy(this.deck, site.slide.id);
  }
}

// PowerPoint.ShapeGroup: the group behind a shape of type Group, and the
// collection of the shapes it holds (a ShapeScopedCollection, keyed by id).
class ShapeGroupProxy extends ShapeBound {
  get id(): string {
    return groupOf(this.deck, this.handleId).id;
  }
  get shapes(): GroupShapeCollectionProxy {
    return new GroupShapeCollectionProxy(this.deck, this.handleId);
  }
}

class GroupShapeCollectionProxy extends ShapeBound {
  get items(): ShapeProxy[] {
    return this.children().map((shape) => new ShapeProxy(this.deck, shape.id));
  }
  getItem(id: string): ShapeProxy {
    return new ShapeProxy(this.deck, member(this.children(), id));
  }
  getItemOrNullObject(id: string): ShapeProxy {
    return new ShapeProxy(this.deck, id, true);
  }
  getCount(): FakeClientResult<number> {
    return new FakeClientResult(this.children().length);
  }

  private children(): FakePptShape[] {
    return groupOf(this.deck, this.handleId).shapes;
  }
}

class TagCollectionProxy extends ShapeBound {
  get items(): TagProxy[] {
    const keys = [...this.model().tags.keys()];
    return keys.map((key) => new TagProxy(this.deck, this.handleId, key));
  }

  // PowerPoint upper-cases tag keys and rejects a value that is not a string.
  add(key: string, value: string): void {
    if (typeof value !== "string") {
      throw new Error(`InvalidArgument: tag "${key}" needs a string value`);
    }
    this.model().tags.set(key.toUpperCase(), value);
  }
  delete(key: string): void {
    this.model().tags.delete(key.toUpperCase());
  }
  getItemOrNullObject(key: string): TagProxy {
    return new TagProxy(this.deck, this.handleId, key.toUpperCase(), true);
  }
}

class TagProxy extends ShapeBound {
  constructor(
    deck: FakePresentation,
    shapeId: string,
    private tagKey: string,
    nullable = false,
  ) {
    super(deck, shapeId, nullable);
  }

  get key(): string {
    return this.nullable && this.read() === undefined ? "" : this.tagKey;
  }
  get value(): string {
    const value = this.read();
    if (value !== undefined) return value;
    if (this.nullable) return "";
    throw gone("tag", this.tagKey);
  }
  get isNullObject(): boolean {
    return this.read() === undefined;
  }

  private read(): string | undefined {
    return this.peek()?.shape.tags.get(this.tagKey);
  }
}
