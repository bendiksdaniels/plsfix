// The deck the fake PowerPoint host serves: slides holding shapes, each shape
// holding the tags link identity lives in, and the moves a user makes on them
// - group, cut and paste, copy, delete. Owns the deck and nothing else; the
// office.js objects that read it live in objects.ts. Invariant: a shape is
// found by id wherever it sits, groups included.

import type { FakeTable } from "./tables";

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

// Every lookup that can fail says so the way PowerPoint does, and objects.ts
// throws the same error when it is asked for a shape that has gone.
export function gone(what: string, id: string): Error {
  const error = new Error(`ItemNotFound: no ${what} "${id}"`);
  return Object.assign(error, { code: "ItemNotFound" });
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
