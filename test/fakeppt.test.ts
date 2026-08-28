import { afterEach, describe, expect, it } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
} from "./fakeppt";

enableStrictLoadSemantics();
afterEach(() => uninstallFakePpt());

describe("fake PowerPoint host", () => {
  it("throws on unloaded reads and serves loaded ones", async () => {
    installFakePpt({ slides: 2 });
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      expect(() => slides.items).toThrow(/PropertyNotLoaded/);
      slides.load("items/id");
      await context.sync();
      expect(slides.items).toHaveLength(2);
      expect(() => slides.items[0]!.shapes.items).toThrow(/PropertyNotLoaded/);
    });
  });

  it("tags upsert with uppercase keys and survive move/copy", async () => {
    const { presentation } = installFakePpt({ slides: 2 });
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      const shape = slides.items[0]!.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        { left: 1, top: 2, width: 3, height: 4 },
      );
      shape.tags.add("smt_link", "a");
      shape.tags.add("SMT_LINK", "b");
      shape.load("id");
      await context.sync();
      const stored = presentation.findShape(shape.id).shape;
      expect([...stored.tags.entries()]).toEqual([["SMT_LINK", "b"]]);
      presentation.moveShape(shape.id, slides.items[1]!.id);
      expect(presentation.slides[1]!.shapes[0]!.tags.get("SMT_LINK")).toBe("b");
      const copy = presentation.copyShape(shape.id, slides.items[0]!.id);
      expect(copy.id).not.toBe(shape.id);
      expect(copy.tags.get("SMT_LINK")).toBe("b");
    });
  });

  it("records fill.setImage and selection-based inserts", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      const shape = slides.items[0]!.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        { left: 0, top: 0, width: 10, height: 5 },
      );
      shape.fill.setImage("PNG1");
      shape.load("id");
      await context.sync();
      expect(presentation.findShape(shape.id).shape.fillImage).toBe("PNG1");
    });
    helpers.selectSlide(presentation.slides[0]!.id);
    await new Promise<void>((resolve) => {
      Office.context.document.setSelectedDataAsync(
        "PNG2",
        {
          coercionType: Office.CoercionType.Image,
          imageLeft: 1,
          imageTop: 2,
          imageWidth: 30,
          imageHeight: 40,
        },
        () => {
          resolve();
        },
      );
    });
    expect(helpers.insertedViaSelection()).toEqual([
      {
        slideId: presentation.slides[0]!.id,
        png: "PNG2",
        box: { left: 1, top: 2, width: 30, height: 40 },
      },
    ]);
    expect(presentation.slides[0]!.shapes.at(-1)!.type).toBe("Image");
  });

  it("reports the slide a shape was moved to", async () => {
    const { presentation } = installFakePpt({ slides: 2 });
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      const shape = slides.items[0]!.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        { left: 0, top: 0, width: 4, height: 3 },
      );
      shape.load("id");
      await context.sync();
      const before = shape.getParentSlideOrNullObject();
      before.load("id,isNullObject");
      await context.sync();
      expect(before.isNullObject).toBe(false);
      expect(before.id).toBe(slides.items[0]!.id);

      presentation.moveShape(shape.id, slides.items[1]!.id);
      const after = shape.getParentSlideOrNullObject();
      after.load("id");
      await context.sync();
      expect(after.id).toBe(slides.items[1]!.id);
    });
  });

  // What a batching change is measured in: the count starts at zero for every
  // installed host, so one test's round trips never land in another's budget.
  it("counts the round trips since it was installed", async () => {
    const { helpers } = installFakePpt({ slides: 1 });
    expect(helpers.syncCount()).toBe(0);
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      await context.sync();
    });
    expect(helpers.syncCount()).toBe(2);
  });

  it("delete removes the shape from the model", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      const shapes = slides.items[0]!.shapes;
      const shape = shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        { left: 0, top: 0, width: 4, height: 3 },
      );
      shape.load("id");
      await context.sync();
      expect(presentation.slides[0]!.shapes).toHaveLength(1);

      shape.delete();
      expect(presentation.slides[0]!.shapes).toHaveLength(0);
      const orphan = shape.getParentSlideOrNullObject();
      orphan.load("isNullObject");
      await context.sync();
      expect(orphan.isNullObject).toBe(true);
    });
  });
});
