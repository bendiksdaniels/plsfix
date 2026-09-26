// ui-window.js: an app window around one 1440 x 900 capture of Excel or PowerPoint for the web
// (the Office header is cropped at capture, so the bar here carries only the file's name).
// Layers are whole shots; crops are one region of a shot in its own box, so a scene can lift,
// slide or reveal just that part. Content px = capture CSS px x k (k = content width / 1440).
window.UiWindow = (() => {
  const BAR = 40;
  const CAP_W = 1440;
  const CAP_H = 900;
  const APPS = { excel: ['#1D8A50', 'pls,fix Demo Model.xlsx'], ppt: ['#D0512E', 'pls,fix Demo Deck.pptx'] };

  function create(parent, cw, app = 'excel') {
    const ch = Math.round(cw * CAP_H / CAP_W);
    const k = cw / CAP_W;
    const el = El.make('div', 'win', parent, { width: `${cw}px`, height: `${ch + BAR}px` });
    const bar = El.make('div', 'win-bar', el);
    for (let i = 0; i < 3; i++) El.make('span', 'win-dot', bar);
    const title = El.make('div', 'win-title', bar);
    El.make('span', 'win-app', title, { background: APPS[app][0] });
    El.make('span', null, title).textContent = APPS[app][1];
    const content = El.make('div', 'win-content', el, { width: `${cw}px`, height: `${ch}px` });

    // a whole shot filling the content
    function layer(name, into) {
      return El.img(Shots.src(name), into || content, { position: 'absolute', left: '0px', top: '0px', width: `${cw}px`, height: `${ch}px` });
    }

    // one region of a shot (capture CSS px, grown by `pad`), in its own positioned box
    function crop(name, r, pad = 0, into) {
      const x = (r.x - pad) * k, y = (r.y - pad) * k;
      const box = El.make('div', 'crop', into || content, { left: `${x}px`, top: `${y}px`, width: `${(r.w + 2 * pad) * k}px`, height: `${(r.h + 2 * pad) * k}px` });
      El.img(Shots.src(name), box, { position: 'absolute', left: `${-x}px`, top: `${-y}px`, width: `${cw}px`, height: `${ch}px` });
      return box;
    }

    // a ring (glow or alarm) around a capture rect, grown by `pad`
    function ring(r, cls = 'glow-ring', pad = 5) {
      return El.make('div', cls, content, { left: `${(r.x - pad) * k}px`, top: `${(r.y - pad) * k}px`, width: `${(r.w + 2 * pad) * k}px`, height: `${(r.h + 2 * pad) * k}px` });
    }

    // a capture point (CSS px) in content px
    const at = (x, y) => ({ x: x * k, y: y * k });
    const mid = (r) => at(r.x + r.w / 2, r.y + r.h / 2);

    return { el, content, layer, crop, ring, at, mid, k, cw, ch, h: ch + BAR, BAR };
  }

  return { create, BAR, CAP_W, CAP_H };
})();
