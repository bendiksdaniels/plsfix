// The deck the fake PowerPoint host serves - slides holding shapes, each shape
// holding the tags link identity lives in - and the office.js objects for
// slides, shapes and tags on top of it. Every object addresses its item by id
// and reads the deck on each access, so it sees a move or a delete happen.

import { FakeClientResult, Loadable } from "./strict";
import { newFakeTable, TableProxy, type FakeTable } from "./tables";

export interface FakePptShape {
  id: string;
  name: string;
  type: "GeometricShape" | "Image" | "Table" | "Group" | "Placeholder";
  left: number;
  top: number;
  width: number;
  height: number;
  zOrder: number;
  tags: Map<string, string>;
  fillImage: string | null;
  lineVisible: boolean;
  setImageCalls: number;
  // PowerPoint.TextFrame.hasText: false on an empty layout placeholder, which
  // is the one shape a new object is allowed to be placed over.
  hasText: boolean;
  // The shapes a group holds; null on everything that is not a group.
  group: FakeShapeGroup | null;
  // The grid a shape of type Table holds; null on everything else.
  table: FakeTable | null;
}

export interface FakeShapeGroup {
  id: string;
  shapes: FakePptShape[];
}

export interface FakeSlide {
  id: string;
  shapes: FakePptShape[];
}

export interface FakeShapeInit {
  name?: string;
  type?: FakePptShape["type"];
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  fillImage?: string;
  hasText?: boolean;
}

export interface ShapeSite {
  slide: FakeSlide;
  shape: FakePptShape;
  // The array holding the shape: the slide's, or a group's inside it.
  siblings: FakePptShape[];
}

// PowerPoint numbers the z-order stack per container, 0 at the bottom.
function renumber(shapes: FakePptShape[]): void {
  shapes.forEach((shape, index) => {
    shape.zOrder = index;
  });
}

// The box a new group takes: the one its members occupy together.
function bounds(shapes: FakePptShape[]): {
  left: number;
  top: number;
  width: number;
  height: number;
} {
  const left = Math.min(...shapes.map((shape) => shape.left));
  const top = Math.min(...shapes.map((shape) => shape.top));
  const right = Math.max(...shapes.map((shape) => shape.left + shape.width));
  const bottom = Math.max(...shapes.map((shape) => shape.top + shape.height));
  return { left, top, width: right - left, height: bottom - top };
}

export class FakePresentation {
  slides: FakeSlide[] = [];
  selectedSlideIds: string[] = [];
  fileUrl = "https://contoso.sharepoint.com/Shared%20Documents/deck.pptx";
  private slideSeq = 0;
  private shapeSeq = 0;

  constructor(slideCount = 1) {
    for (let index = 0; index < slideCount; index += 1) this.addSlide();
  }

  addSlide(): FakeSlide {
    this.slideSeq += 1;
    const slide: FakeSlide = { id: `slide-${this.slideSeq}`, shapes: [] };
    this.slides.push(slide);
    return slide;
  }

  addShape(slide: FakeSlide, init: FakeShapeInit = {}): FakePptShape {
    this.shapeSeq += 1;
    const shape: FakePptShape = {
      id: `shape-${this.shapeSeq}`,
      name: init.name ?? `Shape ${this.shapeSeq}`,
      type: init.type ?? "GeometricShape",
      left: init.left ?? 0,
      top: init.top ?? 0,
      width: init.width ?? 0,
      height: init.height ?? 0,
      zOrder: slide.shapes.length,
      tags: new Map(),
      fillImage: init.fillImage ?? null,
      lineVisible: true,
      setImageCalls: 0,
      hasText: init.hasText ?? false,
      group: null,
      table: null,
    };
    slide.shapes.push(shape);
    renumber(slide.shapes);
    return shape;
  }

  // What Ctrl+G does: the shapes leave the slide - or the group they were in -
  // and live inside a new group shape sized to hold them.
  groupShapes(shapeIds: string[], slideId: string): FakePptShape {
    const slide = this.findSlideOrThrow(slideId);
    const sites = shapeIds.map((id) => this.findShape(id));
    const box = bounds(sites.map((site) => site.shape));
    const children = sites.map((site) => {
      site.siblings.splice(site.siblings.indexOf(site.shape), 1);
      renumber(site.siblings);
      return site.shape;
    });
    renumber(children);
    this.shapeSeq += 1;
    const seq = this.shapeSeq;
    const group: FakePptShape = {
      id: `shape-${seq}`,
      name: `Group ${seq}`,
      type: "Group",
      ...box,
      zOrder: slide.shapes.length,
      tags: new Map(),
      fillImage: null,
      lineVisible: true,
      setImageCalls: 0,
      hasText: false,
      table: null,
      group: { id: `group-${seq}`, shapes: children },
    };
    slide.shapes.push(group);
    renumber(slide.shapes);
    return group;
  }

  // Cut and paste onto another slide: same shape, same tags, same id.
  moveShape(shapeId: string, toSlideId: string): void {
    const { shape, siblings } = this.findShape(shapeId);
    const target = this.findSlideOrThrow(toSlideId);
    siblings.splice(siblings.indexOf(shape), 1);
    target.shapes.push(shape);
    renumber(siblings);
    renumber(target.shapes);
  }

  // Copy and paste: the tags ride along, which is why link ids collide. A
  // group is copied whole, every shape inside it getting its own new id.
  copyShape(shapeId: string, toSlideId: string): FakePptShape {
    const { shape } = this.findShape(shapeId);
    const target = this.findSlideOrThrow(toSlideId);
    const copy = this.cloneShape(shape);
    target.shapes.push(copy);
    renumber(target.shapes);
    return copy;
  }

  deleteShape(shapeId: string): void {
    const site = this.peekShape(shapeId);
    if (!site) return;
    site.siblings.splice(site.siblings.indexOf(site.shape), 1);
    renumber(site.siblings);
  }

  findShape(shapeId: string): ShapeSite {
    const site = this.peekShape(shapeId);
    if (!site) throw gone("shape", shapeId);
    return site;
  }

  // The lookup the objects use: a shape can vanish under them, and one the
  // user grouped is still on its slide - one or more groups down.
  peekShape(shapeId: string): ShapeSite | null {
    for (const slide of this.slides) {
      const site = siteIn(slide, slide.shapes, shapeId);
      if (site) return site;
    }
    return null;
  }

  private cloneShape(shape: FakePptShape): FakePptShape {
    this.shapeSeq += 1;
    const seq = this.shapeSeq;
    const group = shape.group;
    return {
      ...shape,
      id: `shape-${seq}`,
      tags: new Map(shape.tags),
      table: shape.table
        ? (JSON.parse(JSON.stringify(shape.table)) as FakeTable)
        : null,
      group: group
        ? {
            id: `group-${seq}`,
            shapes: group.shapes.map((child) => this.cloneShape(child)),
          }
        : null,
    };
  }

  findSlide(slideId: string): FakeSlide | null {
    return this.slides.find((slide) => slide.id === slideId) ?? null;
  }

  findSlideOrThrow(slideId: string): FakeSlide {
    const slide = this.findSlide(slideId);
    if (!slide) throw gone("slide", slideId);
    return slide;
  }

  selectedSlides(): FakeSlide[] {
    return this.selectedSlideIds
      .map((id) => this.findSlide(id))
      .filter((slide): slide is FakeSlide => slide !== null);
  }
}

// ---------------------------------------------------------------------------
// The objects the host hands out
// ---------------------------------------------------------------------------

function gone(what: string, id: string): Error {
  const error = new Error(`ItemNotFound: no ${what} "${id}"`);
  return Object.assign(error, { code: "ItemNotFound" });
}

// PowerPoint reads a shape group through the group shape, so anything else
// answers the same GeneralException the real host does.
function notAGroup(shapeId: string): Error {
  const error = new Error(`GeneralException: shape "${shapeId}" is no group`);
  return Object.assign(error, { code: "GeneralException" });
}

// Depth-first through the groups: the site names the array holding the shape,
// which is what a delete, a move or a regroup has to splice.
function siteIn(
  slide: FakeSlide,
  siblings: FakePptShape[],
  shapeId: string,
): ShapeSite | null {
  for (const shape of siblings) {
    if (shape.id === shapeId) return { slide, shape, siblings };
    const nested = shape.group
      ? siteIn(slide, shape.group.shapes, shapeId)
      : null;
    if (nested) return nested;
  }
  return null;
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
    options: BoxOptions & { values?: string[][] } = {},
  ): ShapeProxy {
    const slide = this.slide();
    const { values, ...box } = options;
    const shape = this.deck.addShape(slide, {
      name: `Table ${String(slide.shapes.length + 1)}`,
      type: "Table",
      ...box,
    });
    shape.table = newFakeTable(rowCount, columnCount, values);
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
