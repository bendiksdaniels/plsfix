// The deck the fake PowerPoint host serves - slides holding shapes, each shape
// holding the tags link identity lives in - and the office.js objects for
// slides, shapes and tags on top of it. Every object addresses its item by id
// and reads the deck on each access, so it sees a move or a delete happen.

import { FakeClientResult, Loadable } from "./strict";

export interface FakePptShape {
  id: string;
  name: string;
  type: "GeometricShape" | "Image" | "Table";
  left: number;
  top: number;
  width: number;
  height: number;
  zOrder: number;
  tags: Map<string, string>;
  fillImage: string | null;
  lineVisible: boolean;
  setImageCalls: number;
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
}

export interface ShapeSite {
  slide: FakeSlide;
  shape: FakePptShape;
}

// PowerPoint numbers the z-order stack per slide, 0 at the bottom.
function renumber(slide: FakeSlide): void {
  slide.shapes.forEach((shape, index) => {
    shape.zOrder = index;
  });
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
    };
    slide.shapes.push(shape);
    renumber(slide);
    return shape;
  }

  // Cut and paste onto another slide: same shape, same tags, same id.
  moveShape(shapeId: string, toSlideId: string): void {
    const { slide, shape } = this.findShape(shapeId);
    const target = this.findSlideOrThrow(toSlideId);
    slide.shapes.splice(slide.shapes.indexOf(shape), 1);
    target.shapes.push(shape);
    renumber(slide);
    renumber(target);
  }

  // Copy and paste: the tags ride along, which is why link ids collide.
  copyShape(shapeId: string, toSlideId: string): FakePptShape {
    const { shape } = this.findShape(shapeId);
    const target = this.findSlideOrThrow(toSlideId);
    this.shapeSeq += 1;
    const copy: FakePptShape = {
      ...shape,
      id: `shape-${this.shapeSeq}`,
      tags: new Map(shape.tags),
    };
    target.shapes.push(copy);
    renumber(target);
    return copy;
  }

  deleteShape(shapeId: string): void {
    const site = this.peekShape(shapeId);
    if (!site) return;
    site.slide.shapes.splice(site.slide.shapes.indexOf(site.shape), 1);
    renumber(site.slide);
  }

  findShape(shapeId: string): ShapeSite {
    const site = this.peekShape(shapeId);
    if (!site) throw gone("shape", shapeId);
    return site;
  }

  // The lookup the objects use: a shape can vanish under them.
  peekShape(shapeId: string): ShapeSite | null {
    for (const slide of this.slides) {
      const shape = slide.shapes.find((item) => item.id === shapeId);
      if (shape) return { slide, shape };
    }
    return null;
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
    return new ShapeProxy(this.deck, id);
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
  get tags(): TagCollectionProxy {
    return new TagCollectionProxy(this.deck, this.handleId);
  }

  delete(): void {
    this.deck.deleteShape(this.handleId);
  }
  getParentSlideOrNullObject(): SlideProxy {
    return new SlideProxy(this.deck, this.peek()?.slide.id ?? "", true);
  }
}

class ShapeFillProxy extends ShapeBound {
  setImage(base64EncodedImage: string): void {
    const shape = this.model();
    shape.fillImage = base64EncodedImage;
    shape.setImageCalls += 1;
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
