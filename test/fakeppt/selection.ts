// PowerPoint.ShapeScopedCollection (PowerPointApi 1.5): the shapes
// context.presentation.getSelectedShapes() hands back, over the selection ids
// objects.ts keeps for the deck (SlideProxy.setSelectedShapes writes them).
// Owns only the collection and the test-facing selection shortcut. Invariant:
// an id that names no shape can never enter the selection through the
// shortcut either.

import { gone, invalidArgument, type FakePresentation } from "./model";
import { selectedShapeIds, setSelectedShapeIds, ShapeProxy } from "./objects";
import { FakeClientResult, Loadable } from "./strict";

// helpers.selectShapes(ids): sets the selection directly, order preserved,
// skipping the same-slide rule Slide.setSelectedShapes enforces - the same
// shortcut helpers.selectSlide takes against a slide id.
export function selectShapesForTest(
  deck: FakePresentation,
  shapeIds: string[],
): void {
  for (const id of shapeIds) deck.findShape(id);
  setSelectedShapeIds(deck, shapeIds);
}

export class SelectedShapesCollectionProxy extends Loadable {
  constructor(private deck: FakePresentation) {
    super();
  }

  get items(): ShapeProxy[] {
    return selectedShapeIds(this.deck).map(
      (id) => new ShapeProxy(this.deck, id),
    );
  }
  getItem(id: string): ShapeProxy {
    if (!selectedShapeIds(this.deck).includes(id)) throw gone("shape", id);
    return new ShapeProxy(this.deck, id);
  }
  getItemAt(index: number): ShapeProxy {
    const id = selectedShapeIds(this.deck)[index];
    if (id === undefined) throw invalidArgument(`no shape at index ${index}`);
    return new ShapeProxy(this.deck, id);
  }
  getItemOrNullObject(id: string): ShapeProxy {
    const found = selectedShapeIds(this.deck).includes(id);
    return new ShapeProxy(this.deck, id, !found);
  }
  getCount(): FakeClientResult<number> {
    return new FakeClientResult(selectedShapeIds(this.deck).length);
  }
}
