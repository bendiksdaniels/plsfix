// Pure geometry for PowerPoint object tools. The Office adapter loads shapes,
// hands their boxes here, then writes back only the properties each action
// changes. Keeping the layout rules host-free makes the exact result testable.

export interface ObjectBox {
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

export type AlignMode =
  "left" | "center" | "right" | "top" | "middle" | "bottom";

export interface ObjectMove {
  id: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

export function alignBoxes(boxes: ObjectBox[], mode: AlignMode): ObjectMove[] {
  if (boxes.length < 2) return [];
  const left = Math.min(...boxes.map((box) => box.left));
  const top = Math.min(...boxes.map((box) => box.top));
  const right = Math.max(...boxes.map((box) => box.left + box.width));
  const bottom = Math.max(...boxes.map((box) => box.top + box.height));
  const center = (left + right) / 2;
  const middle = (top + bottom) / 2;
  return boxes.map((box) => {
    switch (mode) {
      case "left":
        return { id: box.id, left };
      case "center":
        return { id: box.id, left: center - box.width / 2 };
      case "right":
        return { id: box.id, left: right - box.width };
      case "top":
        return { id: box.id, top };
      case "middle":
        return { id: box.id, top: middle - box.height / 2 };
      case "bottom":
        return { id: box.id, top: bottom - box.height };
    }
  });
}

export function distributeBoxes(
  boxes: ObjectBox[],
  axis: "horizontal" | "vertical",
): ObjectMove[] {
  if (boxes.length < 3) return [];
  const sorted = [...boxes].sort((a, b) =>
    axis === "horizontal" ? a.left - b.left : a.top - b.top,
  );
  const first = sorted[0]!;
  const last = sorted.at(-1)!;
  const start = axis === "horizontal" ? first.left : first.top;
  const end =
    axis === "horizontal" ? last.left + last.width : last.top + last.height;
  const total = sorted.reduce(
    (sum, box) => sum + (axis === "horizontal" ? box.width : box.height),
    0,
  );
  const gap = (end - start - total) / (sorted.length - 1);
  let cursor = start;
  return sorted.map((box) => {
    const move: ObjectMove =
      axis === "horizontal"
        ? { id: box.id, left: cursor }
        : { id: box.id, top: cursor };
    cursor += (axis === "horizontal" ? box.width : box.height) + gap;
    return move;
  });
}

export function matchSize(boxes: ObjectBox[]): ObjectMove[] {
  const source = boxes[0];
  if (!source || boxes.length < 2) return [];
  return boxes.slice(1).map((box) => ({
    id: box.id,
    width: source.width,
    height: source.height,
  }));
}

export function swapBoxes(boxes: ObjectBox[]): ObjectMove[] {
  if (boxes.length !== 2) return [];
  const [first, second] = boxes as [ObjectBox, ObjectBox];
  return [
    { id: first.id, left: second.left, top: second.top },
    { id: second.id, left: first.left, top: first.top },
  ];
}

export function sameKindAndSize(
  source: { type: string; width: number; height: number },
  candidate: { type: string; width: number; height: number },
  tolerance = 1,
): boolean {
  return (
    source.type === candidate.type &&
    Math.abs(source.width - candidate.width) <= tolerance &&
    Math.abs(source.height - candidate.height) <= tolerance
  );
}
