// tools-beats.js: one chapter of the toolbox inside its own Excel window: the "before" shot,
// the pointer's one real click, the "after" shot revealed the way that tool's result reads
// (a wipe across a row, down a block, a chart popping in, a crossfade, the pane scrolling),
// then mint rings on what changed. CHAPTERS is the table; times are relative to the turn.
window.ToolsBeats = (() => {
  const CHAPTERS = [
    { key: 'fmt', before: 'x-fmt', after: 'x-fmt-after', click: ['x-fmt', 'button'], reveal: 'wipeX', area: ['x-fmt', 'row'], rings: [['x-fmt', 'row']] },
    { key: 'charts', before: 'x-bridge', after: 'x-bridge-after', click: ['x-bridge', 'waterfall'], reveal: 'pop', area: ['x-bridge-after', 'chart'], rings: [['x-bridge-after', 'chart']] },
    { key: 'tpl', before: 'x-tpl', after: 'x-tpl-after', click: ['x-tpl', 'dcf'], reveal: 'wipeY', area: ['x-tpl-after', 'block'], rings: [['x-tpl-after', 'block']] },
    { key: 'rec', before: 'x-rec', after: 'x-rec-after', click: ['x-rec', 'find'], reveal: 'fade', area: ['x-rec', 'list'], rings: 'hits' },
    // the hit on the hidden Scratch sheet, and the same words in the bridge table on screen
    { key: 'file', before: 'x-find-before', after: 'x-find', click: ['x-find-before', 'run'], reveal: 'fade', area: ['x-find', 'results'], rings: [['x-find', 'hidden'], ['x-find', 'cell']] },
    // no click: the colours and the number style are ringed on the Brand tab first, then the
    // pane moves on to the shortcuts
    {
      key: 'brand', before: 'x-brand', after: 'x-keys', click: null, reveal: 'fade', area: ['x-keys', 'keys'],
      early: [['x-brand', 'palette'], ['x-brand', 'preview']], rings: [['x-keys', 'keys']],
      R: { early: [0.35, 1.95], reveal: [2.0, 2.35], fade: [2.0, 2.35], rings: [2.5, 4.85] },
    },
  ];
  // early: rings on the before shot, gone before the reveal (a chapter without a click); pop:
  // the new chart springs in lifted, then settles into the shot
  const R = { press: 1.4, early: [0.5, 1.45], reveal: [1.5, 2.25], fade: [1.5, 1.85], rings: [2.3, 4.85], settle: 0.9 };
  const times = (ch) => ({ ...R, ...(ch.R || {}) });

  function build(win, ch, t0) {
    const r = times(ch);
    const s = { ch, t0, r, before: win.layer(ch.before), after: win.layer(ch.after) };
    if (ch.reveal === 'pop') s.pop = win.crop(ch.after, Shots.rect(...ch.area));
    const rects = ch.rings === 'hits'
      ? Shots.keys('x-rec-after').filter((k) => /^hit\d+$/.test(k)).map((k) => Shots.rect('x-rec-after', k)).concat([Shots.rect('x-rec-after', 'target')])
      : ch.rings.map(([shot, key]) => Shots.rect(shot, key));
    s.rings = rects.map((r) => win.ring(r, 'glow-ring', 4));
    s.early = (ch.early || []).map(([shot, key]) => win.ring(Shots.rect(shot, key), 'glow-ring', 4));
    s.early.forEach((_, i) => Stage.cue(t0 + r.early[0] + i * 0.3 + 0.05, 'pop', { pitch: 74 + i * 3, gain: -10 }));
    if (ch.click) Stage.cue(t0 + r.press + 0.02, 'click', { gain: -2 });
    Stage.cue(t0 + r.reveal[0], ch.reveal === 'pop' ? 'pop' : 'whoosh', ch.reveal === 'pop' ? { pitch: 74, gain: -4 } : { dur: ch.reveal === 'fade' ? 0.45 : 0.7, gain: -15 });
    s.rings.forEach((_, i) => Stage.cue(t0 + r.rings[0] + i * 0.12 + 0.05, 'pop', { pitch: 76 + (i % 3) * 3, gain: -10 }));
    return s;
  }

  // where the pointer clicks (capture px), or null for a chapter without a click
  function target(s) {
    if (!s.ch.click) return null;
    const r = Shots.rect(...s.ch.click);
    return { x: r.x + r.w * 0.5, y: r.y + r.h * 0.55 };
  }

  function update(s, win, t) {
    const rt = t - s.t0, r = s.r;
    const span = s.ch.reveal === 'fade' ? r.fade : r.reveal;
    const p = Ease.seg(rt, span[0], span[1], Ease.inOutCubic);
    const on = rt >= r.reveal[0];
    El.show(s.after, 0);
    s.after.style.clipPath = '';
    const kind = s.ch.reveal;
    if (kind === 'fade') El.show(s.after, p);
    else if (kind === 'pop') {
      El.show(s.after, Ease.seg(rt, r.reveal[0], r.reveal[0] + 0.2));
      const sp = Ease.clamp(Ease.spring(rt - r.reveal[0], 2.2, 0.55), 0, 1.15);
      const lift = Math.min(1, sp) * (1 - Ease.seg(rt, r.reveal[0] + r.settle - 0.35, r.reveal[0] + r.settle, Ease.inOutCubic));
      El.show(s.pop, on && rt < r.reveal[0] + r.settle ? 1 : 0);
      El.tf(s.pop, { s: 0.86 + 0.14 * sp, y: (1 - Math.min(1, sp)) * 24 });
      s.pop.style.boxShadow = `0 ${20 * lift}px 44px rgba(0,0,0,${0.35 * lift})`;
    } else wipe(s, win, kind, rt, p, on);
    s.early.forEach((el, i) => {
      El.show(el, Ease.inOut(rt, r.early[0] + i * 0.3, r.early[1], 0.3, 0.2));
      El.tf(el, { s: 1 + 0.2 * (1 - Ease.seg(rt, r.early[0] + i * 0.3, r.early[0] + i * 0.3 + 0.4, Ease.outBack)) });
    });
    s.rings.forEach((el, i) => {
      const a = Ease.inOut(rt, r.rings[0] + i * 0.12, r.rings[1], 0.3, 0.35);
      const beat = 0.5 + 0.5 * Math.cos(2 * Math.PI * (t - s.t0 - r.rings[0]) / T.BEAT / 2);
      El.show(el, a * (0.7 + 0.3 * beat));
      El.tf(el, { s: 1 + 0.2 * (1 - Ease.seg(rt, r.rings[0] + i * 0.12, r.rings[0] + i * 0.12 + 0.4, Ease.outBack)) });
    });
  }

  // the after shot revealed across (wipeX) or down (wipeY) the chapter's area, then whole
  function wipe(s, win, kind, rt, p, on) {
    const a = Shots.rect(...s.ch.area);
    El.show(s.after, on ? 1 : 0);
    if (p >= 1) return;
    const k = win.k;
    const right = kind === 'wipeX' ? win.cw - (a.x + a.w * p) * k : win.cw - (a.x + a.w) * k;
    const bottom = kind === 'wipeY' ? win.ch - (a.y + a.h * p) * k : win.ch - (a.y + a.h) * k;
    s.after.style.clipPath = `inset(${a.y * k}px ${right}px ${bottom}px ${a.x * k}px)`;
  }

  // the subject the camera leans on during the hold (capture px)
  function subject(s) {
    const r = Shots.rect(...s.ch.area);
    return { x: r.x + r.w / 2, y: r.y + r.h / 2 };
  }

  return { CHAPTERS, R, build, target, update, subject };
})();
