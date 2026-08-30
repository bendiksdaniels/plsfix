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
});
