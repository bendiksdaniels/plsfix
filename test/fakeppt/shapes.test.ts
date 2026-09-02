// The shapes a chart is drawn from, against the fake host: text boxes, lines,
// groups made from ids, a pie's adjustments (the one object PowerPoint refuses
// to hand out before the shape's first sync) and the formats each carries.
// Strict load semantics are on, so every read here is loaded and synced first.

import { afterEach, describe, expect, it } from "vitest";
import { enableStrictLoadSemantics, installFakePpt, uninstallFakePpt } from ".";

enableStrictLoadSemantics();
afterEach(() => {
  uninstallFakePpt();
});

function box(
  left: number,
  top: number,
  width: number,
  height: number,
): { left: number; top: number; width: number; height: number } {
  return { left, top, width, height };
}

describe("fake PowerPoint shapes", () => {
  it("groups shapes by id, tags the group and deletes the members with it", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const a = shapes.addGeometricShape("Rectangle", box(0, 0, 10, 10));
      const b = shapes.addTextBox("1 234", box(20, 0, 30, 18));
      a.load("id");
      b.load("id");
      await c.sync();
      const g = shapes.addGroup([a.id, b.id]);
      g.name = "pls,fix chart x";
      g.tags.add("PLSFIX_LINK", "t");
      g.load("id,type,left,top,width,height");
      await c.sync();
      expect(g.type).toBe("Group");
      expect([g.left, g.top, g.width, g.height]).toEqual([0, 0, 50, 18]);
      const all = c.presentation.slides.getItemAt(0).shapes;
      all.load("items/type");
      await c.sync();
      expect(all.items.map((s) => s.type)).toEqual(["Group"]);
      // The members left the slide's top level for the group's own shapes.
      const members = presentation.findShape(g.id).shape.group!.shapes;
      expect(members.map((s) => s.type)).toEqual(["GeometricShape", "TextBox"]);
      g.delete();
      await c.sync();
      const after = c.presentation.slides.getItemAt(0).shapes;
      after.load("items/id");
      await c.sync();
      expect(after.items).toHaveLength(0);
      expect(presentation.peekShape(members[0]!.id)).toBeNull();
    });
  });

  it("refuses an empty group", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      expect(() => shapes.addGroup([])).toThrow(/InvalidArgument/);
    });
  });

  it("refuses a negative width or height, on add and on a later write", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      expect(() =>
        shapes.addGeometricShape("Rectangle", box(0, 0, -5, 10)),
      ).toThrow(/InvalidArgument/);
      expect(() => shapes.addLine("Straight", box(0, 0, 10, -5))).toThrow(
        /InvalidArgument/,
      );
      const rect = shapes.addGeometricShape("Rectangle", box(0, 0, 10, 10));
      expect(() => {
        rect.width = -1;
      }).toThrow(/InvalidArgument/);
      expect(() => {
        rect.height = -1;
      }).toThrow(/InvalidArgument/);
    });
  });

  it("shapes a pie only after its first sync and normalises the angles", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const pie = c.presentation.slides
        .getItemAt(0)
        .shapes.addGeometricShape("Pie", box(0, 0, 100, 100));
      expect(() => {
        pie.adjustments.set(0, -90);
      }).toThrow(/InvalidParam/);
      await c.sync();
      pie.adjustments.set(0, -90);
      pie.adjustments.set(1, 200);
      await c.sync();
      const a = pie.adjustments.get(0);
      const b = pie.adjustments.get(1);
      pie.adjustments.load("count");
      await c.sync();
      expect([pie.adjustments.count, a.value, b.value]).toEqual([2, -90, -160]);
    });
  });

  it("gives a shape with no adjustment point none to set", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const rect = c.presentation.slides
        .getItemAt(0)
        .shapes.addGeometricShape("Rectangle", box(0, 0, 10, 10));
      rect.adjustments.load("count");
      await c.sync();
      expect(rect.adjustments.count).toBe(0);
      expect(() => {
        rect.adjustments.set(0, 12);
      }).toThrow(/InvalidArgument/);
    });
  });

  it("records text formatting, fills and line formats", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const bar = shapes.addGeometricShape("Rectangle", box(0, 0, 40, 60));
      bar.fill.setSolidColor("#B27E54");
      bar.lineFormat.visible = false;
      const label = shapes.addTextBox("1 234", box(0, 62, 47, 18));
      label.textFrame.wordWrap = false;
      label.textFrame.textRange.font.name = "Aptos Narrow";
      label.textFrame.textRange.font.size = 9;
      label.textFrame.textRange.font.color = "#282623";
      label.textFrame.textRange.font.bold = true;
      label.textFrame.textRange.paragraphFormat.horizontalAlignment = "Center";
      const rule = shapes.addLine("Straight", box(0, 61, 200, 0));
      rule.lineFormat.color = "#282623";
      rule.lineFormat.weight = 0.75;
      bar.load("id");
      label.load("id");
      rule.load("id");
      await c.sync();

      expect(presentation.slides[0]!.shapes.map((s) => s.type)).toEqual([
        "GeometricShape",
        "TextBox",
        "Line",
      ]);
      expect(presentation.findShape(bar.id).shape).toMatchObject({
        geometry: "Rectangle",
        fillColor: "#B27E54",
        fillCleared: false,
        lineVisible: false,
      });
      expect(presentation.findShape(label.id).shape).toMatchObject({
        text: "1 234",
        hasText: true,
        wordWrap: false,
        alignment: "Center",
        font: { name: "Aptos Narrow", size: 9, color: "#282623", bold: true },
      });
      expect(presentation.findShape(rule.id).shape).toMatchObject({
        geometry: "Straight",
        lineColor: "#282623",
        lineWeight: 0.75,
      });

      bar.fill.clear();
      expect(presentation.findShape(bar.id).shape).toMatchObject({
        fillColor: null,
        fillCleared: true,
      });
    });
  });

  it("is a Mac until a test says the host is something else", () => {
    const { helpers } = installFakePpt({ slides: 1 });
    expect(Office.context.platform).toBe("Mac");
    helpers.setPlatform("OfficeOnline");
    expect(Office.context.platform).toBe(Office.PlatformType.OfficeOnline);
  });

  it("derives fill.type, reads foregroundColor, and reads/writes fill.transparency", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const rect = shapes.addGeometricShape("Rectangle", box(0, 0, 10, 10));
      rect.fill.setSolidColor("#B27E54");
      rect.fill.load("type,foregroundColor,transparency");
      await c.sync();
      expect(rect.fill.type).toBe("Solid");
      expect(rect.fill.foregroundColor).toBe("#B27E54");
      expect(rect.fill.transparency).toBe(0);
      rect.fill.transparency = 0.5;
      rect.fill.clear();
      rect.fill.load("type,transparency");
      await c.sync();
      expect(rect.fill.type).toBe("NoFill");
      // PowerPointApi 1.4: null once the fill type no longer supports one.
      expect(rect.fill.transparency).toBeNull();
    });
  });

  it("gives a picture fill the type PictureAndTexture", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const pic = shapes.addGeometricShape("Rectangle", box(0, 0, 10, 10));
      pic.fill.setImage("<png>");
      pic.fill.load("type");
      await c.sync();
      expect(pic.fill.type).toBe("PictureAndTexture");
    });
  });

  it("reads and writes line.transparency, dashStyle and style, defaulted like PowerPoint's own", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const rect = shapes.addGeometricShape("Rectangle", box(0, 0, 10, 10));
      rect.lineFormat.load("transparency,dashStyle,style");
      await c.sync();
      expect(rect.lineFormat.transparency).toBe(0);
      expect(rect.lineFormat.dashStyle).toBe("Solid");
      expect(rect.lineFormat.style).toBe("Single");
      rect.lineFormat.transparency = 0.25;
      rect.lineFormat.dashStyle = "DashDot";
      rect.lineFormat.style = "ThickThin";
      rect.lineFormat.load("transparency,dashStyle,style");
      await c.sync();
      expect(rect.lineFormat.transparency).toBe(0.25);
      expect(rect.lineFormat.dashStyle).toBe("DashDot");
      expect(rect.lineFormat.style).toBe("ThickThin");
    });
  });

  it("nulls weight, dashStyle and style once the line is invisible", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const rect = shapes.addGeometricShape("Rectangle", box(0, 0, 10, 10));
      rect.lineFormat.weight = 2;
      rect.lineFormat.dashStyle = "DashDot";
      rect.lineFormat.style = "ThickThin";
      rect.lineFormat.visible = false;
      rect.lineFormat.load("weight,dashStyle,style");
      await c.sync();
      // PowerPointApi 1.4: null while the line isn't visible, whatever was
      // last written - a capture that skips visibility loses these for good.
      expect(rect.lineFormat.weight).toBeNull();
      expect(rect.lineFormat.dashStyle).toBeNull();
      expect(rect.lineFormat.style).toBeNull();
      expect(() => {
        rect.lineFormat.weight = null as unknown as number;
      }).toThrow(/InvalidArgument/);
    });
  });

  it("throws PropertyNotLoaded for an unloaded fill or line scalar", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const rect = shapes.addGeometricShape("Rectangle", box(0, 0, 10, 10));
      rect.load("id");
      await c.sync();
      expect(() => rect.fill.type).toThrow(/PropertyNotLoaded/);
      expect(() => rect.lineFormat.dashStyle).toThrow(/PropertyNotLoaded/);
    });
  });

  it("returns the selection in the order it was made, with getCount", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 1 });
    const slide = presentation.slides[0]!;
    const a = presentation.addShape(slide, box(0, 0, 10, 10));
    const b = presentation.addShape(slide, box(20, 0, 10, 10));
    helpers.selectShapes([b.id, a.id]);
    await PowerPoint.run(async (c) => {
      const selected = c.presentation.getSelectedShapes();
      selected.load("items/id");
      const count = selected.getCount();
      await c.sync();
      expect(selected.items.map((s) => s.id)).toEqual([b.id, a.id]);
      expect(count.value).toBe(2);
    });
  });

  it("addresses the selection by index and by id, and refuses one outside it", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 1 });
    const slide = presentation.slides[0]!;
    const a = presentation.addShape(slide, box(0, 0, 10, 10));
    const b = presentation.addShape(slide, box(20, 0, 10, 10));
    const outside = presentation.addShape(slide, box(40, 0, 10, 10));
    helpers.selectShapes([b.id, a.id]);
    await PowerPoint.run(async (c) => {
      const selected = c.presentation.getSelectedShapes();
      const first = selected.getItemAt(0);
      const byId = selected.getItem(a.id);
      first.load("id");
      byId.load("id");
      await c.sync();
      expect(first.id).toBe(b.id);
      expect(byId.id).toBe(a.id);
      expect(() => selected.getItem(outside.id)).toThrow(/ItemNotFound/);
      expect(() => selected.getItemAt(5)).toThrow(/InvalidArgument/);
    });
    expect(() => helpers.selectShapes(["bogus"])).toThrow(/ItemNotFound/);
  });

  it("Slide.setSelectedShapes refuses a shape from elsewhere, else replaces the selection", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 2 });
    const [slide0, slide1] = presentation.slides as [
      (typeof presentation.slides)[0],
      (typeof presentation.slides)[0],
    ];
    const a = presentation.addShape(slide0, box(0, 0, 10, 10));
    const b = presentation.addShape(slide0, box(20, 0, 10, 10));
    const elsewhere = presentation.addShape(slide1, box(0, 0, 10, 10));
    helpers.selectSlide(slide1.id);
    await PowerPoint.run(async (c) => {
      const slide = c.presentation.slides.getItemAt(0);
      expect(() => {
        slide.setSelectedShapes([a.id, elsewhere.id]);
      }).toThrow(/InvalidArgument/);
      slide.setSelectedShapes([b.id, a.id]);
    });
    expect(presentation.selectedSlideIds).toEqual([slide0.id]);
    await PowerPoint.run(async (c) => {
      const selected = c.presentation.getSelectedShapes();
      selected.load("items/id");
      await c.sync();
      expect(selected.items.map((s) => s.id)).toEqual([b.id, a.id]);
    });
  });

  it("getParentSlide finds the shape's slide, or throws once the shape is gone", async () => {
    const { presentation } = installFakePpt({ slides: 2 });
    const slide1 = presentation.slides[1]!;
    const shape = presentation.addShape(slide1, box(0, 0, 10, 10));
    await PowerPoint.run(async (c) => {
      const handle = c.presentation.slides
        .getItem(slide1.id)
        .shapes.getItem(shape.id);
      const parent = handle.getParentSlide();
      parent.load("id");
      await c.sync();
      expect(parent.id).toBe(slide1.id);
      handle.delete();
      await c.sync();
      expect(() => handle.getParentSlide()).toThrow(/ItemNotFound/);
    });
  });
});
