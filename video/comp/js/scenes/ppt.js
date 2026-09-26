// ppt.js (60.0 - 80.4 s): "and yes, PowerPoint too". Excel and PowerPoint windows arrive side
// by side under the title; the camera turns to Excel for "Export active chart", the chart flies
// across into PowerPoint's Inbox, Insert puts it on the slide as shapes; a growth assumption
// changes, Push all; PowerPoint's Links tab says "Update available", Update all repaints the
// slide while the camera pushes in on it; the Tools tab lines three shapes up.
(() => {
  const B = PptBeats.K;
  const P = (o) => Object.freeze(o);
  const SPLIT = { x: P({ x: -470, y: 90, s: 0.6, rx: 0, ry: 0 }), p: P({ x: 470, y: 90, s: 0.6, rx: 0, ry: 0 }) };
  const FOCUS_X = { x: P({ x: -30, y: 20, s: 0.94, rx: 2, ry: -4 }), p: P({ x: 980, y: 60, s: 0.5, rx: 0, ry: -24 }) };
  const FOCUS_P = { x: P({ x: -980, y: 60, s: 0.5, rx: 0, ry: 24 }), p: P({ x: 30, y: 20, s: 0.94, rx: 2, ry: 4 }) };
  const OFF = { x: P({ x: -1800, y: 90, s: 0.6, rx: 0, ry: 30 }), p: P({ x: 1800, y: 90, s: 0.6, rx: 0, ry: -30 }) };
  const AWAY = { x: P({ x: -700, y: 700, s: 0.3, rx: 20, ry: 0 }), p: P({ x: 700, y: 700, s: 0.3, rx: 20, ry: 0 }) };
  // [t, layout] the camera moves to over the next 0.6 s
  const MOVES = [[60.2, SPLIT], [62.4, FOCUS_X], [63.9, SPLIT], [65.3, FOCUS_P], [68.8, FOCUS_X], [72.75, FOCUS_P], [79.6, AWAY]];
  const FLY = [64.35, 65.3];
  const K = { title: [60.3, 62.3], caps: [['ppt.export', 62.6, 65.9], ['ppt.native', 66.3, 68.65], ['ppt.update', 69.2, 76.6], ['ppt.tools', 77.0, 80.1]] };
  // the push-in on the slide chart during Update all: its scale and where the chart lands
  const LEAN = { s: 2.3, aim: { x: 960, y: 470 } };

  Stage.scene({
    id: 'ppt', from: T.ppt.from, to: T.ppt.to,
    build(root) {
      this.x = this.window(root, 'excel', PptBeats.buildExcel);
      this.p = this.window(root, 'ppt', PptBeats.buildPpt);
      this.clicks = PptBeats.clicks();
      const chart = Shots.rect('x-link', 'chart');
      this.card = this.x.win.crop('x-link', chart, 0, root);
      this.card.classList.add('lift');
      Object.assign(this.card.style, { left: '0px', top: '0px', transformOrigin: '0 0', zIndex: '6' });
      this.title = UiType.headline(root, Copy.t('ppt.title'), { left: '0px', right: '0px', top: '90px', textAlign: 'center', fontSize: '84px' });
      Stage.copy(this.title.el, 61.6, 'ppt.title');
      this.caps = K.caps.map(([key, a, b]) => {
        const c = UiType.caption(root, Copy.t(key));
        Stage.copy(c.el, (a + b) / 2, key);
        return { c, a, b };
      });
      Stage.cue(60.0, 'section', { name: 'break' });
      Stage.cue(64.0, 'section', { name: 'launch' });
      Stage.cue(60.2, 'whoosh', { dur: 0.9, pan: -0.6, gain: -8 });
      Stage.cue(60.3, 'whoosh', { dur: 0.9, pan: 0.6, gain: -8 });
      MOVES.slice(1).forEach(([t]) => Stage.cue(t, 'whoosh', { dur: 0.6, gain: -9 }));
      Stage.cue(FLY[0], 'riser', { dur: FLY[1] - FLY[0], gain: -9 });
      Stage.cue(FLY[1], 'tap', { gain: -3 });
    },
    window(root, app, build) {
      const win = UiWindow.create(root, Cam.HERO.cw, app);
      const at = Cam.box(win, Cam.HERO.cx, Cam.HERO.cy);
      Cam.place(win, at);
      const state = build(win);
      return { win, at, state, cursor: UiPointer.cursor(win.content), ring: UiPointer.ring(win.content), key: app === 'excel' ? 'x' : 'p' };
    },
    pose(side, t) {
      let pose = t < MOVES[0][0] + 0.9 ? Cam.mix(OFF[side], SPLIT[side], Ease.seg(t, MOVES[0][0], MOVES[0][0] + 0.9, Ease.outExpo)) : { ...SPLIT[side] };
      for (const [at, lay] of MOVES.slice(1)) {
        if (t < at) break;
        pose = Cam.mix(pose, lay[side], Ease.seg(t, at, at + 0.6, at === 79.6 ? Ease.inCubic : Ease.outExpo));
      }
      // a slow sway, so a hold never stands still (none during the chart's flight: its path
      // is computed from the resting poses)
      const sway = 1 - Ease.inOut(t, FLY[0] - 0.5, FLY[1] + 0.3, 0.3, 0.3, Ease.inOutCubic, Ease.inOutCubic);
      return { ...pose, ry: pose.ry + sway * 1.6 * Math.sin((t - 60) * 0.9), y: pose.y + sway * 5 * Math.sin((t - 60) * 0.6) };
    },
    update(t) {
      this.title.update(t, K.title[0], K.title[1]);
      this.caps.forEach(({ c, a, b }) => c.update(t, a, b));
      for (const w of [this.x, this.p]) {
        const pose = w.key === 'p' ? this.lean(w, this.pose('p', t), t) : this.pose(w.key, t);
        El.tf(w.win.el, pose);
        El.show(w.win.el, 1 - Ease.p(t, 80.0, 80.35));
        w.win.el.style.zIndex = pose.s > 0.7 ? '3' : '1';
      }
      PptBeats.updateExcel(this.x.state, this.x.win, t);
      PptBeats.updatePpt(this.p.state, this.p.win, t);
      this.pointer(this.x, this.clicks.excel, t);
      this.pointer(this.p, this.clicks.ppt, t);
      this.fly(t);
    },
    // PowerPoint leans in on the slide chart, flat and framed, while Update all repaints it
    lean(w, pose, t) {
      const L = PptBeats.lean(t);
      if (L.w <= 0) return pose;
      const f = Cam.fit(w.win, w.at, LEAN.s, Cam.focus(w.win, w.at, w.win.at(L.x, L.y), LEAN.s, LEAN.aim));
      return Cam.mix(pose, { ...pose, s: LEAN.s, x: f.x, y: f.y, rx: 0, ry: 0 }, L.w);
    },
    // the exported chart as a card, lifted off Excel's sheet and dropped on PowerPoint's Inbox
    fly(t) {
      const q = Ease.seg(t, FLY[0], FLY[1], Ease.inOutCubic);
      const on = t >= FLY[0] - 0.15 && t < FLY[1] + 0.12;
      El.show(this.card, on ? 1 - Ease.p(t, FLY[1], FLY[1] + 0.12) : 0);
      if (!on) return;
      const from = this.stage(this.x, SPLIT.x, Shots.rect('x-link', 'chart'));
      const to = this.stage(this.p, SPLIT.p, Shots.rect('p-inbox', 'row'));
      const lift = Ease.seg(t, FLY[0] - 0.15, FLY[0] + 0.1);
      const x = Ease.lerp(from.x, to.x, q), y = Ease.lerp(from.y, to.y, q) - 160 * Math.sin(Math.PI * q);
      const s = Ease.lerp(from.s, to.s * 0.34, q) * (1 + 0.06 * lift * (1 - q));
      El.tf(this.card, { x, y, s, rz: -3 * Math.sin(Math.PI * q) });
    },
    // a capture rect's top-left on the stage and its scale, for a window resting at `pose`
    stage(w, pose, r) {
      const cx = w.at.x + w.win.cw / 2, cy = w.at.y + w.win.h / 2;
      const px = w.at.x + r.x * w.win.k, py = w.at.y + w.win.BAR + r.y * w.win.k;
      return { x: cx + pose.s * (px - cx) + pose.x, y: cy + pose.s * (py - cy) + pose.y, s: pose.s };
    },
    // clicks under CHAIN s apart are one continuous path (the pointer glides from the last
    // press to the next); a lone click comes in from below right and leaves after the press
    pointer(w, list, t) {
      const CHAIN = 2.0;
      const i = list.findIndex((c, j) => {
        const prev = list[j - 1], next = list[j + 1];
        const from = prev && c.press - prev.press < CHAIN ? prev.press + 0.15 : c.press - 1.0;
        const to = next && next.press - c.press < CHAIN ? c.press + 0.15 : c.press + 0.6;
        return t >= from && t < to;
      });
      if (i < 0) {
        w.cursor.set(0, 0, 0, 0);
        w.ring.set(0, 0, 0);
        return;
      }
      const c = list[i], prev = list[i - 1], next = list[i + 1];
      const chained = prev && c.press - prev.press < CHAIN;
      const start = chained ? prev : { x: c.x + 200, y: c.y + 240 };
      const q = Ease.seg(t, chained ? prev.press + 0.15 : c.press - 0.9, c.press - 0.1, Ease.inOutCubic);
      const pt = w.win.at(Ease.lerp(start.x, c.x, q), Ease.lerp(start.y, c.y, q));
      const press = Ease.inOut(t, c.press - 0.08, c.press + 0.12, 0.06, 0.1);
      const fadeIn = chained ? 1 : Ease.p(t, c.press - 1.0, c.press - 0.8);
      const fadeOut = next && next.press - c.press < CHAIN ? 1 : 1 - Ease.p(t, c.press + 0.4, c.press + 0.6);
      w.cursor.set(pt.x, pt.y, press, fadeIn * fadeOut);
      const r = w.win.at(c.x, c.y);
      const life = Ease.p(t, c.press - 0.05, c.press + 0.45);
      w.ring.set(r.x, r.y, life > 0 && life < 1 ? life : 0);
    },
  });
})();
