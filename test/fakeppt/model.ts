// The deck the fake PowerPoint host serves: slides holding shapes, each shape
// holding the tags link identity lives in, and the moves a user makes on them
// - group, cut and paste, copy, delete. Owns the deck and nothing else; the
// office.js objects that read it live in objects.ts. Invariant: a shape is
// found by id wherever it sits, groups included.

import type { FakeTable } from "./tables";

// The insets a text frame keeps, null wherever the add-in left PowerPoint's own.
export interface FakeShapeMargins {
  left: number | null;
  right: number | null;
  top: number | null;
  bottom: number | null;
}

export interface FakeShapeFont {
  name?: string;
  size?: number;
  color?: string;
  bold?: boolean;
}

export interface FakePptShape {
  id: string;
  name: string;
  type:
    | "GeometricShape"
    | "Image"
    | "Table"
    | "Group"
    | "Placeholder"
    | "TextBox"
    | "Line";
  // The geometry the shape was drawn with, "Rectangle" or "Pie"; a line keeps
  // its connector type here, and a shape with neither keeps null.
  geometry: string | null;
  left: number;
  top: number;
  width: number;
  height: number;
  zOrder: number;
  tags: Map<string, string>;
  fillImage: string | null;
  fillColor: string | null;
  fillCleared: boolean;
  lineVisible: boolean;
  lineColor: string | null;
  lineWeight: number | null;
  // A pie's [start, end] in degrees; empty on a geometry with no adjustment.
  adjustments: number[];
  setImageCalls: number;
  // PowerPoint.TextFrame.hasText: false on an empty layout placeholder, which
  // is the one shape a new object is allowed to be placed over.
  hasText: boolean;
  text: string | null;
  font: FakeShapeFont;
  // The paragraph's horizontal alignment; the frame's own settings sit beside.
  alignment: string | null;
  verticalAlignment: string | null;
  autoSize: string | null;
  wordWrap: boolean | null;
  margins: FakeShapeMargins;
  // False until the sync after the add: PowerPoint hands out a shape's
  // adjustments only once the host has heard of the shape.
  synced: boolean;
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
  geometry?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  fillImage?: string;
  hasText?: boolean;
  text?: string;
}

// PowerPoint gives every geometry a fixed number of adjustment points. The
// fake draws pies alone, whose two are the start and the end angle; the
// defaults are the quarter PowerPoint itself starts a pie with.
const ADJUSTMENTS: Readonly<Record<string, readonly number[]>> = {
  Pie: [0, -90],
};

function adjustmentsFor(geometry: string | null): number[] {
  return [...(ADJUSTMENTS[geometry ?? ""] ?? [])];
}

// A fresh shape carries no format of its own: what the add-in writes is what
// the deck records, so a property nobody set reads as null.
function blankFormat(): Pick<
  FakePptShape,
  | "fillColor"
  | "fillCleared"
  | "lineColor"
  | "lineWeight"
  | "text"
  | "font"
  | "alignment"
  | "verticalAlignment"
  | "autoSize"
  | "wordWrap"
  | "margins"
> {
  return {
    fillColor: null,
    fillCleared: false,
    lineColor: null,
    lineWeight: null,
    text: null,
    font: {},
    alignment: null,
    verticalAlignment: null,
    autoSize: null,
    wordWrap: null,
    margins: { left: null, right: null, top: null, bottom: null },
  };
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
    const geometry = init.geometry ?? null;
    const shape: FakePptShape = {
      ...blankFormat(),
      id: `shape-${this.shapeSeq}`,
      name: init.name ?? `Shape ${this.shapeSeq}`,
      type: init.type ?? "GeometricShape",
      geometry,
      left: init.left ?? 0,
      top: init.top ?? 0,
      width: init.width ?? 0,
      height: init.height ?? 0,
      zOrder: slide.shapes.length,
      tags: new Map(),
      fillImage: init.fillImage ?? null,
      lineVisible: true,
      adjustments: adjustmentsFor(geometry),
      setImageCalls: 0,
      hasText: init.hasText ?? (init.text ?? "").length > 0,
      synced: false,
      group: null,
      table: null,
      text: init.text ?? null,
    };
    slide.shapes.push(shape);
    renumber(slide.shapes);
    return shape;
  }

  // What Ctrl+G does: the shapes leave the slide - or the group they were in -
  // and live inside a new group shape sized to hold them.
  groupShapes(shapeIds: string[], slideId: string): FakePptShape {
    if (shapeIds.length === 0) throw invalidArgument("a group needs a shape");
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
      ...blankFormat(),
      id: `shape-${seq}`,
      name: `Group ${seq}`,
      type: "Group",
      geometry: null,
      ...box,
      zOrder: slide.shapes.length,
      tags: new Map(),
      fillImage: null,
      lineVisible: true,
      adjustments: [],
      setImageCalls: 0,
      hasText: false,
      synced: false,
      table: null,
      group: { id: `group-${seq}`, shapes: children },
    };
    slide.shapes.push(group);
    renumber(slide.shapes);
    return group;
  }

  // What a context.sync() does to a batch of adds: every shape the deck holds
  // has been round-tripped once, which is when PowerPoint starts handing out
  // its adjustments. Called by the runtime, never by a proxy.
  markSynced(): void {
    const walk = (shapes: FakePptShape[]): void => {
      for (const shape of shapes) {
        shape.synced = true;
        if (shape.group) walk(shape.group.shapes);
      }
    };
    for (const slide of this.slides) walk(slide.shapes);
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
      adjustments: [...shape.adjustments],
      font: { ...shape.font },
      margins: { ...shape.margins },
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

// Every lookup that can fail says so the way PowerPoint does, and objects.ts
// throws the same error when it is asked for a shape that has gone.
export function gone(what: string, id: string): Error {
  const error = new Error(`ItemNotFound: no ${what} "${id}"`);
  return Object.assign(error, { code: "ItemNotFound" });
}

// The argument PowerPoint refuses outright: an empty group, an adjustment
// index the geometry does not have.
export function invalidArgument(why: string): Error {
  const error = new Error(`InvalidArgument: ${why}`);
  return Object.assign(error, { code: "InvalidArgument" });
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
