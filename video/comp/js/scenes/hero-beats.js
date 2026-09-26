// hero-beats.js: what happens inside the Excel hero window, beat by beat, from the real shots:
// Autocolor recolours the model block (a wipe across it), Audit overlay scans the numbers top to
// bottom and the planted hardcode in F16 is ringed, Precedents rings D14 and D16 and draws the
// arrows into D17, then the Workbook tab: Run model check is pressed and its list lifts row by row.
window.HeroBeats = (() => {
  const K = {
    color: { press: 9.9, wipe: [10.0, 10.8] },
    audit: { press: 12.9, scan: [13.0, 13.8], ring: [14.0, 16.7] },
    trace: { cell: 16.95, press: 17.95, swap: [18.0, 18.25], rings: [18.35, 20.9], arrows: [18.5, 19.3] },
    check: { before: [20.95, 21.15], press: 21.45, scroll: [21.65, 22.05], rows: 22.1, rowGap: 0.16, out: 25.0 },
  };

  const PANE_BG = '#EEF2F5'; // the task pane's own ground, under a scrolled shot

  function build(win) {
    const s = { base: win.layer('x-pnl') };
    s.color = win.layer('x-pnl-color');
    s.audit = win.layer('x-pnl-audit');
    s.scan = El.make('div', 'fill', win.content, { background: 'linear-gradient(180deg, rgba(46,196,182,0) 0%, rgba(46,196,182,0.55) 85%, rgba(46,196,182,0.95) 100%)', inset: 'auto', left: '0px', width: `${win.cw}px`, height: '70px' });
    s.alarm = win.ring(Shots.rect('x-pnl', 'f16'), 'alarm-ring', 3);
    s.trace = win.layer('x-pnl-trace');
    s.subject = win.ring(Shots.rect('x-pnl', 'd17'), 'glow-ring', 3);
    s.rings = ['d14', 'd16'].map((k) => win.ring(Shots.rect('x-pnl', k), 'glow-ring', 3));
    s.arrows = arrows(win, ['d14', 'd16'].map((k) => Shots.rect('x-pnl', k)), Shots.rect('x-pnl', 'd17'));
    s.before = win.layer('x-check-before');
    s.check = win.layer('x-check');
    // the two Model check shots sit at different pane scroll positions: the pane scrolls up to
    // the findings as they arrive, both shots moving together inside the pane's box
    const pane = Shots.rect('x-pnl', 'pane');
    s.dy = Shots.rect('x-check-before', 'run').y - Shots.rect('x-check', 'run').y;
    s.scrollBefore = win.crop('x-check-before', pane);
    s.scrollBefore.style.background = PANE_BG;
    s.scrollAfter = win.crop('x-check', pane);
    s.rows = Shots.rects('x-check', 'rows').slice(0, 7).map((r) => {
      const el = win.crop('x-check', r);
      el.style.borderRadius = `${10 * win.k}px`;
      return el;
    });
    cues();
    return s;
  }

  function cues() {
    Stage.cue(K.color.press + 0.02, 'click', { gain: -2 });
    Stage.cue(K.color.wipe[0], 'whoosh', { dur: 0.8, pan: 0.3, gain: -14 });
    Stage.cue(K.audit.press + 0.02, 'click', { gain: -2 });
    Stage.cue(K.audit.scan[0], 'whoosh', { dur: 0.8, gain: -13 });
    Stage.cue(K.audit.ring[0] + 0.05, 'pop', { pitch: 62, gain: -6 });
    Stage.cue(K.trace.cell + 0.02, 'click', { gain: -4 });
    Stage.cue(K.trace.press + 0.02, 'click', { gain: -2 });
    [0, 1].forEach((i) => Stage.cue(K.trace.rings[0] + i * 0.15 + 0.05, 'pop', { pitch: 79 + i * 4, gain: -8 }));
    Stage.cue(K.check.before[0], 'whoosh', { dur: 0.35, gain: -17 });
    Stage.cue(K.check.press + 0.02, 'click', { gain: -2 });
    Stage.cue(K.check.scroll[0], 'whoosh', { dur: 0.4, gain: -16 });
    for (let i = 0; i < 7; i++) Stage.cue(K.check.rows + i * K.check.rowGap + 0.05, 'pop', { pitch: 72 + (i % 4) * 3, gain: -11 });
  }

  // curved mint arrows from each precedent into the traced cell (Excel's own trace arrows, in
  // the video's colour), drawn from the source out
  function arrows(win, from, to) {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('class', 'arrow');
    svg.setAttribute('width', win.cw);
    svg.setAttribute('height', win.ch);
    win.content.appendChild(svg);
    const end = win.at(to.x + to.w * 0.3, to.y + to.h * 0.5);
    return from.map((r) => {
      const a = win.at(r.x + r.w * 0.3, r.y + r.h * 0.5);
      const c = { x: a.x - 60 * win.k, y: (a.y + end.y) / 2 };
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      path.setAttribute('d', `M ${a.x} ${a.y} Q ${c.x} ${c.y} ${end.x} ${end.y}`);
      Object.assign(path.style, { fill: 'none', stroke: '#2EC4B6', strokeWidth: '3', strokeLinecap: 'round' });
      svg.appendChild(path);
      const dot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      dot.setAttribute('cx', a.x);
      dot.setAttribute('cy', a.y);
      dot.setAttribute('r', 5);
      dot.style.fill = '#2EC4B6';
      svg.appendChild(dot);
      const len = path.getTotalLength();
      path.style.strokeDasharray = `${len}`;
      return { path, dot, len };
    });
  }

  function update(s, win, t) {
    colour(s, win, t);
    audit(s, win, t);
    trace(s, win, t);
    check(s, win, t);
  }

  // Autocolor: the recoloured shot is revealed across the model block, then whole (its toast)
  function colour(s, win, t) {
    const b = Shots.rect('x-pnl', 'block');
    const p = Ease.seg(t, K.color.wipe[0], K.color.wipe[1], Ease.inOutCubic);
    El.show(s.color, t >= K.color.wipe[0] ? 1 : 0);
    const whole = t >= K.color.wipe[1] + 0.1;
    const right = win.cw - (b.x + b.w * p) * win.k;
    s.color.style.clipPath = whole ? '' : `inset(${b.y * win.k}px ${right}px ${win.ch - (b.y + b.h) * win.k}px ${b.x * win.k}px)`;
  }

  // Audit overlay: the audited shot scans in under a mint line; F16 gets the alarm ring
  function audit(s, win, t) {
    const p = Ease.seg(t, K.audit.scan[0], K.audit.scan[1], Ease.inOutCubic);
    El.show(s.audit, t >= K.audit.scan[0] ? 1 : 0);
    s.audit.style.clipPath = p >= 1 ? '' : `inset(0 0 ${(1 - p) * win.ch}px 0)`;
    El.show(s.scan, p > 0 && p < 1 ? 1 : 0);
    El.tf(s.scan, { y: p * win.ch - 70 });
    const on = Ease.inOut(t, K.audit.ring[0], K.audit.ring[1], 0.3, 0.3);
    const beat = 0.5 + 0.5 * Math.cos(2 * Math.PI * (t - K.audit.ring[0]) / T.BEAT / 2);
    El.show(s.alarm, on * (0.65 + 0.35 * beat));
    El.tf(s.alarm, { s: 1 + 0.12 * (1 - beat) * on });
  }

  function trace(s, win, t) {
    El.show(s.trace, Ease.seg(t, K.trace.swap[0], K.trace.swap[1]));
    const subj = Ease.inOut(t, K.trace.cell, K.trace.rings[1], 0.25, 0.3);
    El.show(s.subject, subj);
    s.rings.forEach((el, i) => {
      const a = Ease.inOut(t, K.trace.rings[0] + i * 0.15, K.trace.rings[1], 0.3, 0.3);
      El.show(el, a);
      El.tf(el, { s: 1 + 0.25 * (1 - Ease.seg(t, K.trace.rings[0] + i * 0.15, K.trace.rings[0] + i * 0.15 + 0.4, Ease.outBack)) });
    });
    const fade = 1 - Ease.seg(t, K.trace.rings[1] - 0.3, K.trace.rings[1]);
    s.arrows.forEach((a, i) => {
      const d = Ease.seg(t, K.trace.arrows[0] + i * 0.12, K.trace.arrows[1] + i * 0.12, Ease.inOutCubic);
      a.path.style.strokeDashoffset = `${a.len * (1 - d)}`;
      El.show(a.path, d > 0 ? fade : 0);
      El.show(a.dot, d > 0 ? fade : 0);
    });
  }

  // Model check: the Workbook tab with its button, the press, then the list, its rows lifting
  // one by one with a mint edge
  function check(s, win, t) {
    El.show(s.before, Ease.seg(t, K.check.before[0], K.check.before[1]));
    const e = Ease.seg(t, K.check.scroll[0], K.check.scroll[1], Ease.inOutCubic);
    const scrolling = t >= K.check.scroll[0] && t < K.check.scroll[1];
    El.show(s.scrollBefore, scrolling ? 1 : 0);
    El.show(s.scrollAfter, scrolling ? e : 0);
    s.scrollBefore.firstChild.style.transform = `translateY(${-e * s.dy * win.k}px)`;
    s.scrollAfter.firstChild.style.transform = `translateY(${(1 - e) * s.dy * win.k}px)`;
    El.show(s.check, t >= K.check.scroll[1] ? 1 : 0);
    s.rows.forEach((el, i) => {
      const t0 = K.check.rows + i * K.check.rowGap;
      const p = Ease.clamp(Ease.spring(t - t0, 2.2, 0.6), 0, 1.15) * (1 - Ease.seg(t, K.check.out - 0.3, K.check.out));
      El.show(el, t >= t0 && t < K.check.out ? 1 : 0);
      El.tf(el, { x: -10 * p, s: 1 + 0.03 * p });
      el.style.boxShadow = `0 ${14 * p}px ${30 * p}px rgba(0,0,0,${0.35 * Math.min(1, p)}), 0 0 0 ${2 * Math.min(1, p)}px rgba(46,196,182,${Math.min(1, p)})`;
    });
  }

  return { build, update, K };
})();
