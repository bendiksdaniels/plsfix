// ui-pointer.js: the mouse pointer (a macOS-style arrow) and its click ring.
// Each returns {el, set(...)}; coordinates are in the parent's own px, so a pointer placed
// inside a window content box moves with the window's 3D transform.
window.UiPointer = (() => {
  const ARROW = '<svg viewBox="0 0 28 40" width="28" height="40" aria-hidden="true"><path d="M2 2 L2 31 L9.2 24.3 L14 36 L19 34 L14.3 22.6 L24 22.6 Z" fill="#fff" stroke="#111" stroke-width="1.8" stroke-linejoin="round"/></svg>';

  function cursor(parent) {
    const el = El.make('div', 'cursor', parent);
    el.innerHTML = ARROW;
    // press 0..1 squeezes the arrow a little, as a click reads on screen
    function set(x, y, press = 0, opacity = 1) {
      El.tf(el, { x, y, s: 1 - 0.12 * press });
      El.show(el, opacity);
    }
    return { el, set };
  }

  // a gold ring that grows and fades at a click point; p = 0..1 through its life
  function ring(parent) {
    const el = El.make('div', 'click-ring', parent);
    function set(x, y, p) {
      const on = p > 0 && p < 1;
      El.tf(el, { x: x - 30, y: y - 30, s: 0.3 + 1.1 * Ease.outCubic(p) });
      El.show(el, on ? 0.9 * (1 - Ease.inCubic(p)) : 0);
    }
    return { el, set };
  }

  return { cursor, ring };
})();
