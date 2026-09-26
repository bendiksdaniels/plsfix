// cam.js: camera math and the layout slots scenes hand windows between (the hero rest and the
// toolbox grid). A window's box sits at
// (x, y) untransformed and scales about its own centre; focus() gives the translation that puts
// one content point at a chosen stage point.
window.Cam = (() => {
  // where a hero window rests, and the toolbox's contact sheet: two rows of three slots
  const HERO = { cw: 1400, cx: 960, cy: 500 };
  const GRID = { s: 0.34, xs: [420, 960, 1500], ys: [470, 815] };

  // the pose (relative to HERO) that puts a window into grid slot i (row-major)
  const slot = (i) => Object.freeze({ s: GRID.s, x: GRID.xs[i % 3] - HERO.cx, y: GRID.ys[Math.floor(i / 3)] - HERO.cy, rx: 0, ry: 0 });

  // box of a window of content width cw centred at (cx, cy)
  function box(win, cx, cy) {
    return { x: cx - win.cw / 2, y: cy - win.h / 2 };
  }

  // translation for scale s that moves content point p to stage point f
  function focus(win, at, p, s, f) {
    const ox = at.x + win.cw / 2, oy = at.y + win.h / 2;
    const px = at.x + p.x, py = at.y + win.BAR + p.y;
    return { x: f.x - ox - s * (px - ox), y: f.y - oy - s * (py - oy) };
  }

  // a lean's translation t, kept framed: on an axis where the scaled window is larger than the
  // stage it still covers the stage (no empty ground beside the subject), where it is smaller it
  // stays inside with a margin (never cut off); the tilt's few degrees are left out
  const MARGIN = 24;
  function fit(win, at, s, t) {
    const axis = (o, size, stage, v) => {
      const half = (s * size) / 2;
      const lo = half >= stage / 2 ? stage - half : half + MARGIN;
      const hi = half >= stage / 2 ? half : stage - half - MARGIN;
      const c = lo <= hi ? Math.min(hi, Math.max(lo, o + v)) : stage / 2;
      return c - o;
    };
    return { x: axis(at.x + win.cw / 2, win.cw, Stage.W, t.x), y: axis(at.y + win.h / 2, win.h, Stage.H, t.y) };
  }

  // places a window's box once (motion is then only transform)
  function place(win, at) {
    win.el.style.left = `${at.x}px`;
    win.el.style.top = `${at.y}px`;
  }

  // blend two plain objects of numbers
  function mix(a, b, q) {
    const out = {};
    for (const k of Object.keys(a)) out[k] = Ease.lerp(a[k], b[k] != null ? b[k] : a[k], q);
    return out;
  }

  // directional motion blur from this frame's movement (stage px per 1/60 s), through the SVG
  // filter `id` in index.html; a slow lean (under WHIP px a frame) stays sharp and unfiltered
  const WHIP = 10;
  function blur(el, id, dx, dy) {
    const sx = Math.min(36, Math.max(0, Math.abs(dx) - WHIP) * 0.35), sy = Math.min(36, Math.max(0, Math.abs(dy) - WHIP) * 0.35);
    if (sx < 0.5 && sy < 0.5) {
      el.style.filter = '';
      return;
    }
    document.querySelector(`#${id} feGaussianBlur`).setAttribute('stdDeviation', `${sx.toFixed(2)} ${sy.toFixed(2)}`);
    el.style.filter = `url(#${id})`;
  }

  return { HERO, GRID, slot, box, focus, fit, place, mix, blur };
})();
