// ppt-beats.js: what the two windows of the PowerPoint act show, from the real shots. Excel:
// the Links tab before and after "Export active chart", the growth assumption before and after
// the edit, the Links tab before and after Push all. PowerPoint: the empty Inbox, the export
// arriving in it, the chart inserted on slide 3, the Links list with "Update available", the
// chart after Update all (its bars rising in), and three loose shapes the Tools tab lines up.
window.PptBeats = (() => {
  const K = {
    exportPress: 63.4, exportAfter: 63.5, arrive: 65.25,
    insertPress: 66.1, insertAfter: 66.2, settle: 66.95, chartRing: [66.5, 68.6],
    assume: 69.1, edit: 69.9, cellRing: [69.2, 70.8], pushBefore: 70.85, pushPress: 71.4, pushAfter: 71.5, pushRing: [71.65, 72.8],
    linksPress: 73.35, links: 73.45, statusRing: [73.6, 74.5], updatePress: 74.25, grow: [74.35, 74.95], updRing: [75.0, 76.3],
    lean: [74.2, 74.9, 76.3, 76.85],
    tools: 76.95, alignPress: 77.7, align: [77.8, 78.3], distPress: 78.55, dist: [78.65, 79.2],
  };

  // a layer that fades in over the one before it at `at`; with `rise`, it comes in from the
  // bottom of that rect up (bars growing), then whole
  function swap(win, shot, at, list, rise) {
    const el = win.layer(shot);
    list.push({ el, at, rise, win });
    return el;
  }

  function buildExcel(win) {
    const s = { steps: [] };
    win.layer('x-link');
    swap(win, 'x-link-after', K.exportAfter, s.steps);
    swap(win, 'x-assume', K.assume, s.steps);
    swap(win, 'x-assume-after', K.edit, s.steps);
    swap(win, 'x-push-before', K.pushBefore, s.steps);
    swap(win, 'x-push', K.pushAfter, s.steps);
    s.cell = win.ring(Shots.rect('x-assume', 'cell'), 'glow-ring', 4);
    s.chart = win.ring(Shots.rect('x-push', 'chart'), 'glow-ring', 4);
    Stage.cue(K.exportPress + 0.02, 'click', { gain: -2 });
    Stage.cue(K.edit, 'tap', { gain: -5 });
    Stage.cue(K.edit + 0.05, 'pop', { pitch: 84, gain: -8 });
    Stage.cue(K.pushPress + 0.02, 'click', { gain: -2 });
    return s;
  }

  function buildPpt(win) {
    const s = { steps: [] };
    win.layer('p-empty');
    swap(win, 'p-inbox', K.arrive, s.steps);
    swap(win, 'p-inserted', K.insertAfter, s.steps);
    swap(win, 'p-links', K.links, s.steps);
    swap(win, 'p-updated', K.grow[0], s.steps, { rect: chartSpan(), until: K.grow[1] });
    swap(win, 'p-tools', K.tools, s.steps);
    s.chartNew = win.ring(Shots.rect('p-inserted', 'chart'), 'glow-ring', 5);
    s.chartUpd = win.ring(Shots.rect('p-updated', 'chart'), 'glow-ring', 5);
    s.status = win.ring(Shots.rect('p-links', 'status'), 'glow-ring', 4);
    s.pop = win.crop('p-inserted', Shots.rect('p-inserted', 'chart'), 4);
    // the three shapes: cut from the loose shot, moved to where Align and Distribute put them
    const keys = Shots.keys('p-tools').filter((k) => /^shape\d$/.test(k));
    // the empty white slide where each shape sat, so the moving copies leave nothing behind
    s.hide = keys.map((k) => {
      const r = Shots.rect('p-tools', k), pad = 9;
      return El.make('div', 'crop', win.content, { left: `${(r.x - pad) * win.k}px`, top: `${(r.y - pad) * win.k}px`, width: `${(r.w + 2 * pad) * win.k}px`, height: `${(r.h + 2 * pad) * win.k}px`, background: '#FFFFFF' });
    });
    s.shapes = keys.map((k) => ({ el: win.crop('p-tools', Shots.rect('p-tools', k), 3), from: Shots.rect('p-tools', k), to: Shots.rect('p-tools-after', k) }));
    s.after = swap(win, 'p-tools-after', K.dist[1] + 0.05, s.steps);
    Stage.cue(K.insertPress + 0.02, 'click', { gain: -2 });
    Stage.cue(K.insertAfter + 0.05, 'pop', { pitch: 76, gain: -4 });
    Stage.cue(K.linksPress + 0.02, 'click', { gain: -3 });
    Stage.cue(K.statusRing[0] + 0.05, 'pop', { pitch: 79, gain: -9 });
    Stage.cue(K.updatePress + 0.02, 'click', { gain: -2 });
    Stage.cue(K.grow[0], 'riser', { dur: K.grow[1] - K.grow[0], gain: -10 });
    Stage.cue(K.grow[1], 'pop', { pitch: 81, gain: -6 });
    Stage.cue(K.alignPress + 0.02, 'click', { gain: -3 });
    Stage.cue(K.align[0], 'whoosh', { dur: 0.5, gain: -16 });
    Stage.cue(K.distPress + 0.02, 'click', { gain: -3 });
    Stage.cue(K.dist[0], 'whoosh', { dur: 0.55, pan: 0.3, gain: -16 });
    return s;
  }

  function steps(s, t) {
    for (const st of s.steps) {
      if (!st.rise) {
        El.show(st.el, Ease.seg(t, st.at, st.at + 0.22));
        continue;
      }
      const r = st.rise.rect, k = st.win.k, p = Ease.seg(t, st.at, st.rise.until, Ease.inOutCubic);
      El.show(st.el, t >= st.at ? 1 : 0);
      st.el.style.clipPath = t >= st.rise.until + 0.15 ? ''
        : `inset(${(r.y + r.h * (1 - p)) * k}px ${st.win.cw - (r.x + r.w) * k}px ${st.win.ch - (r.y + r.h) * k}px ${r.x * k}px)`;
    }
  }

  // the slide chart before and after Update all (a redrawn chart may be a little larger)
  function chartSpan() {
    const a = Shots.rect('p-links', 'chart'), b = Shots.rect('p-updated', 'chart');
    const x = Math.min(a.x, b.x), y = Math.min(a.y, b.y);
    return { x, y, w: Math.max(a.x + a.w, b.x + b.w) - x, h: Math.max(a.y + a.h, b.y + b.h) - y };
  }

  // the PowerPoint window's push-in on the slide chart while it updates: [weight, centre]
  function lean(t) {
    const w = Ease.seg(t, K.lean[0], K.lean[1], Ease.inOutCubic) * (1 - Ease.seg(t, K.lean[2], K.lean[3], Ease.inOutCubic));
    const c = chartSpan();
    return { w, x: c.x + c.w / 2, y: c.y + c.h / 2 };
  }

  function pulse(el, t, span, amp = 0.16) {
    const a = Ease.inOut(t, span[0], span[1], 0.3, 0.35);
    const beat = 0.5 + 0.5 * Math.cos(2 * Math.PI * (t - span[0]) / T.BEAT / 2);
    El.show(el, a * (0.7 + 0.3 * beat));
    El.tf(el, { s: 1 + amp * (1 - Ease.seg(t, span[0], span[0] + 0.4, Ease.outBack)) });
  }

  function updateExcel(s, win, t) {
    steps(s, t);
    pulse(s.cell, t, K.cellRing);
    pulse(s.chart, t, K.pushRing, 0.06);
  }

  function updatePpt(s, win, t) {
    steps(s, t);
    // the inserted chart springs in, then settles flat into the slide (a lifted card would read
    // as a pasted picture)
    const sp = Ease.clamp(Ease.spring(t - K.insertAfter, 2.3, 0.55), 0, 1.15);
    const lift = Math.min(1, sp) * (1 - Ease.seg(t, K.settle - 0.35, K.settle, Ease.inOutCubic));
    El.show(s.pop, t >= K.insertAfter && t < K.settle ? 1 : 0);
    El.tf(s.pop, { s: 0.85 + 0.15 * sp });
    s.pop.style.boxShadow = `0 ${18 * lift}px 40px rgba(0,0,0,${0.3 * lift})`;
    pulse(s.chartNew, t, K.chartRing, 0.06);
    pulse(s.status, t, K.statusRing, 0.12);
    pulse(s.chartUpd, t, K.updRing, 0.04);
    const moving = t >= K.tools + 0.25 && t < K.dist[1] + 0.05;
    s.hide.forEach((el) => El.show(el, moving ? 1 : 0));
    const a = Ease.seg(t, K.align[0], K.align[1], Ease.inOutCubic);
    const d = Ease.seg(t, K.dist[0], K.dist[1], Ease.inOutCubic);
    s.shapes.forEach(({ el, from, to }) => {
      El.show(el, moving ? 1 : 0);
      El.tf(el, { x: (to.x - from.x) * d * win.k, y: (to.y - from.y) * a * win.k });
    });
  }

  // where each window's pointer clicks (capture px), in order
  function clicks() {
    const c = (shot, key, fx = 0.5, fy = 0.55) => { const r = Shots.rect(shot, key); return { x: r.x + r.w * fx, y: r.y + r.h * fy }; };
    return {
      excel: [{ ...c('x-link', 'export'), press: K.exportPress }, { ...c('x-push-before', 'push'), press: K.pushPress }],
      ppt: [{ ...c('p-inbox', 'insert'), press: K.insertPress }, { ...c('p-inserted', 'linksTab'), press: K.linksPress },
        { ...c('p-links', 'update'), press: K.updatePress },
        { ...c('p-tools', 'align'), press: K.alignPress }, { ...c('p-tools', 'distribute'), press: K.distPress }],
    };
  }

  return { K, buildExcel, buildPpt, updateExcel, updatePpt, clicks, lean };
})();
