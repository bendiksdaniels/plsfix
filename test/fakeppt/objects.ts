// The office.js objects the fake PowerPoint host hands out: slides, shapes,
// their fills, lines, text frames and tags, over the deck in model.ts. Owns
// nothing itself. Invariant: every object addresses its item by id and reads
// the deck on each access, so it sees a move or a delete happen under it.

import {
  gone,
  type FakePptShape,
  type FakePresentation,
  type FakeShapeGroup,
  type FakeSlide,
  type ShapeSite,
} from "./model";
import { FakeClientResult, Loadable } from "./strict";
import { newFakeTable, TableProxy } from "./tables";

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
    const slide = this.slide();
    const shape = this.deck.addShape(slide, {
      // PowerPoint names a new shape after its geometry, "Rectangle 3".
      name: `${geometry} ${String(slide.shapes.length + 1)}`,
      type: "GeometricShape",
      ...options,
    });
    return new ShapeProxy(this.deck, shape.id);
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

class ShapeProxy extends ShapeBound {
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
    this.model().width = value;
  }
  get height(): number {
    return this.model().height;
  }
  set height(value: number) {
    this.model().height = value;
  }

  get fill(): ShapeFillProxy {
    return new ShapeFillProxy(this.deck, this.handleId);
  }
  get lineFormat(): ShapeLineProxy {
    return new ShapeLineProxy(this.deck, this.handleId);
  }
  get textFrame(): ShapeTextFrameProxy {
    return new ShapeTextFrameProxy(this.deck, this.handleId);
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

class ShapeFillProxy extends ShapeBound {
  setImage(base64EncodedImage: string): void {
    const shape = this.model();
    shape.fillImage = base64EncodedImage;
    shape.setImageCalls += 1;
  }
}

// Only what placement reads: whether the shape holds any text, which is how
// an empty layout placeholder is told from an object somebody put there.
class ShapeTextFrameProxy extends ShapeBound {
  get hasText(): boolean {
    return this.model().hasText;
  }
}

class ShapeLineProxy extends ShapeBound {
  get visible(): boolean {
    return this.model().lineVisible;
  }
  set visible(value: boolean) {
    this.model().lineVisible = value;
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
