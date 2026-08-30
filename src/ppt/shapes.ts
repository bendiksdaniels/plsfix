// Walking the shapes of a deck, groups included: PowerPointApi 1.8 exposes
// Shape.group.shapes, so a picture the user dragged into a group is still
// reachable - by the path of group ids running from the slide down to it, and
// that path is how the same shape is addressed again on a refresh.

export const SHAPE_PROPERTIES =
  "items/id,items/type,items/left,items/top,items/width,items/height";

// Groups (Shape.type "Group", Shape.group, ShapeGroup.shapes) and
// fill.setImage both arrived in PowerPointApi 1.8, so one version gates both.
export const GROUP_API = "1.8";
// PowerPoint.ShapeType.group; the typings have no TextBox member, so a shape's
// type is compared as the string the host reports.
export const GROUP_TYPE = "Group";

// A user can nest groups without limit; three levels is every real deck, and
// the cap is what keeps a pathological one from costing a sync per level.
const MAX_DEPTH = 3;

// Where a shape sits: on a slide, or inside groups nested in one.
export interface ShapePath {
  slideId: string;
  shapeId: string;
  groupPath?: string[];
}

export interface PlacedShape {
  slideId: string;
  slideIndex: number;
  shape: PowerPoint.Shape;
  groupPath: string[];
}

// A host that reports no requirements at all is the newest one, the web.
export function hasPowerPointApi(version: string): boolean {
  const requirements = Office.context?.requirements;
  return requirements
    ? requirements.isSetSupported("PowerPointApi", version)
    : true;
}

export function isGrouped(path: ShapePath): boolean {
  return (path.groupPath?.length ?? 0) > 0;
}

// The shape a link lives on, addressed the way it was found: down through its
// groups when it has any, straight off the slide when it has none.
export function shapeAt(
  context: PowerPoint.RequestContext,
  path: ShapePath,
): PowerPoint.Shape {
  const shapes = context.presentation.slides.getItem(path.slideId).shapes;
  const groups = path.groupPath ?? [];
  const outermost = groups[0];
  if (outermost === undefined) return shapes.getItem(path.shapeId);
  let shape = shapes.getItem(outermost);
  for (const groupId of groups.slice(1)) {
    shape = shape.group.shapes.getItem(groupId);
  }
  return shape.group.shapes.getItem(path.shapeId);
}

// Every group at one nesting level is opened in a single batch, so depth costs
// a round trip and width costs none. Below 1.8 there is nothing to open.
export async function expandGroups(
  context: PowerPoint.RequestContext,
  top: PlacedShape[],
): Promise<PlacedShape[]> {
  if (!hasPowerPointApi(GROUP_API)) return top;
  const all = [...top];
  let frontier = top;
  for (let depth = 0; depth < MAX_DEPTH && frontier.length > 0; depth += 1) {
    frontier = await openGroups(context, frontier);
    all.push(...frontier);
  }
  return all;
}

// The children of every group in the frontier, each carrying the path that
// reaches it: its parent's path plus the id of the group it came out of.
async function openGroups(
  context: PowerPoint.RequestContext,
  frontier: PlacedShape[],
): Promise<PlacedShape[]> {
  const opened = frontier
    .filter((entry) => entry.shape.type === GROUP_TYPE)
    .map((entry) => {
      const shapes = entry.shape.group.shapes;
      shapes.load(SHAPE_PROPERTIES);
      return { entry, shapes };
    });
  if (opened.length === 0) return [];
  await context.sync();
  return opened.flatMap(({ entry, shapes }) =>
    shapes.items.map((shape) => ({
      slideId: entry.slideId,
      slideIndex: entry.slideIndex,
      shape,
      groupPath: [...entry.groupPath, entry.shape.id],
    })),
  );
}
