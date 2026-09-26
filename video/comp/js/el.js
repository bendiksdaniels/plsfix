// el.js: small DOM helpers the scenes build with: make an element, set styles, set a 3D
// transform from named parts, show an element at an opacity. No layout logic lives here.
window.El = (() => {
  function make(tag, cls, parent, css) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (css) Object.assign(e.style, css);
    if (parent) parent.appendChild(e);
    return e;
  }

  function img(src, parent, css) {
    const e = make('img', null, parent, css);
    // sync: a shot Chrome dropped from its decode cache is decoded again before the frame paints,
    // never shown blank or low-res in a screenshot (30 shots at 2880x1800 outgrow the cache)
    e.decoding = 'sync';
    e.src = src;
    e.draggable = false;
    return e;
  }

  // transform from parts: x y z (px), s | sx sy (scale), rx ry rz (deg); order: translate, rotate, scale
  function tf(e, o) {
    const parts = [];
    if (o.x || o.y || o.z) parts.push(`translate3d(${o.x || 0}px, ${o.y || 0}px, ${o.z || 0}px)`);
    if (o.rx) parts.push(`rotateX(${o.rx}deg)`);
    if (o.ry) parts.push(`rotateY(${o.ry}deg)`);
    if (o.rz) parts.push(`rotateZ(${o.rz}deg)`);
    const sx = o.sx != null ? o.sx : o.s != null ? o.s : 1;
    const sy = o.sy != null ? o.sy : o.s != null ? o.s : 1;
    if (sx !== 1 || sy !== 1) parts.push(`scale(${sx}, ${sy})`);
    e.style.transform = parts.join(' ');
  }

  // opacity, hiding the element entirely at 0 so it costs nothing to paint; when shown it
  // inherits visibility, so a hidden parent still hides it
  function show(e, o) {
    const v = Math.max(0, Math.min(1, o));
    e.style.opacity = String(v);
    e.style.visibility = v <= 0.001 ? 'hidden' : '';
  }

  return { make, img, tf, show };
})();
