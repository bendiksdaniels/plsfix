// hero.js (6.0 - 26.4 s): "tools for financial models, right in Excel". The intro's mint lifts
// off a floating Excel window (the demo P&L, the pls,fix ribbon tab and pane); the pointer runs
// Autocolor, Audit overlay, Precedents and the Model check (hero-beats.js) while the camera
// leans in on each result; then the window drops into the toolbox's first grid slot.
(() => {
  const B = HeroBeats.K;
  const K = {
    mint: [6.4, 6.95], arrive: [6.4, 7.9], title: [6.65, 8.35], rise: [8.3, 9.05],
    pushAudit: [13.4, 14.4], pushTrace: [16.3, 17.2], pushCheck: [21.65, 22.5], handoff: [25.15, 26.25],
    caps: [['hero.autocolor', 9.3, 12.45], ['hero.audit', 12.85, 16.6], ['hero.trace', 17.05, 20.95], ['hero.check', 21.5, 25.0]],
  };
  const REST = Object.freeze({ s: 1, rx: 4, ry: -6, x: 0, y: 0 });
  const AIM = { x: 900, y: 500 };

  Stage.scene({
    id: 'hero', from: T.hero.from, to: T.hero.to,
    build(root) {
      this.win = UiWindow.create(root, Cam.HERO.cw, 'excel');
      this.at = Cam.box(this.win, Cam.HERO.cx, Cam.HERO.cy);
      Cam.place(this.win, this.at);
      this.beats = HeroBeats.build(this.win);
      this.cursor = UiPointer.cursor(this.win.content);
      this.ring = UiPointer.ring(this.win.content);
      this.path = this.pointerPath();
      this.mint = El.make('div', 'layer', root, { background: 'var(--mint)' });
      this.title = UiType.headline(root, Copy.t('hero.title'), { left: '0px', right: '0px', top: '66px', textAlign: 'center', fontSize: '76px' });
      this.caps = K.caps.map(([key, a, b]) => {
        const c = UiType.caption(root, Copy.t(key));
        Stage.copy(c.el, (a + b) / 2, key);
        return { c, a, b };
      });
      Stage.copy(this.title.el, 7.9, 'hero.title');
      Stage.cue(K.mint[0], 'section', { name: 'groove' });
      [['title', 6.9, 9.3], ['autocolor', 9.4, 12.65], ['audit', 12.9, 16.85], ['trace', 17.1, 21.2], ['check', 21.5, 25.3]]
        .forEach(([k, a, b]) => Stage.voice(`vo.hero.${k}`, a, b));
      Stage.cue(K.mint[0], 'whoosh', { dur: 0.55, gain: -5 });
      Stage.cue(K.rise[0], 'whoosh', { dur: 0.75, gain: -12 });
      Stage.cue(K.handoff[0], 'whoosh', { dur: 1.1, pan: -0.4, gain: -7 });
    },
    // the pointer's stops (capture px): arrive by `at`, press at `press`
    pointerPath() {
      const c = (r, fx = 0.5, fy = 0.5) => ({ x: r.x + r.w * fx, y: r.y + r.h * fy });
      return [
        { ...c(Shots.rect('x-pnl', 'autocolor'), 0.55), at: 9.8, press: B.color.press },
        { ...c(Shots.rect('x-pnl', 'audit'), 0.55), at: 12.8, press: B.audit.press },
        { ...c(Shots.rect('x-pnl', 'd17'), 0.6), at: 16.85, press: B.trace.cell },
        { ...c(Shots.rect('x-pnl', 'precedents'), 0.55), at: 17.85, press: B.trace.press },
        { ...c(Shots.rect('x-check-before', 'run'), 0.3), at: 21.35, press: B.check.press },
      ];
    },
    update(t) {
      const w = this.win;
      const g = Ease.seg(t, K.mint[0], K.mint[1], Ease.outCubic);
      El.show(this.mint, t < K.mint[0] - 0.01 ? 0 : 1 - Ease.p(t, K.mint[1] - 0.05, K.mint[1]));
      El.tf(this.mint, { y: -1080 * g });
      this.title.update(t, K.title[0], K.title[1]);
      this.caps.forEach(({ c, a, b }) => c.update(t, a, b));
      const pose = this.camera(t);
      El.tf(w.el, pose);
      const prev = this.camera(t - 1 / 60);
      Cam.blur(w.el, 'mb0', pose.x - prev.x, pose.y - prev.y);
      El.show(w.el, t >= K.mint[0] ? 1 : 0);
      HeroBeats.update(this.beats, w, t);
      this.pointer(t);
    },
    // the window's pose: low under the title, up to rest with a slow turn, a lean on each
    // beat's subject, then down into the toolbox grid's first slot
    camera(t) {
      const arrive = Ease.seg(t, K.arrive[0], K.arrive[1], Ease.outExpo);
      const start = { s: 0.8, rx: 18, ry: -20, x: 0, y: 230 };
      const low = { s: 0.72, rx: 8, ry: -8, x: 0, y: 160 };
      let pose = Cam.mix(Cam.mix(start, low, arrive), REST, Ease.seg(t, K.rise[0], K.rise[1], Ease.inOutCubic));
      pose.ry += 3 * Ease.p(t, K.rise[1], 25);
      const lean = this.lean(t);
      if (lean.w > 0) {
        const f = Cam.fit(this.win, this.at, lean.s, Cam.focus(this.win, this.at, this.win.at(lean.x, lean.y), lean.s, AIM));
        pose = Cam.mix(pose, { ...pose, s: lean.s, x: f.x, y: f.y }, lean.w);
      }
      return Cam.mix(pose, Cam.slot(0), Ease.seg(t, K.handoff[0], K.handoff[1], Ease.inOutExpo));
    },
    // which subject the camera leans on now: [weight, x, y (capture px), scale]
    lean(t) {
      const f16 = Shots.rect('x-pnl', 'f16'), d15 = Shots.rect('x-pnl', 'd14');
      const list = Shots.rect('x-check', 'list');
      const keys = [
        [K.pushAudit[0], 0, f16.x, f16.y, 1], [K.pushAudit[1], 1, f16.x, f16.y + 20, 1.3],
        [K.pushTrace[0], 1, f16.x, f16.y + 20, 1.3], [K.pushTrace[1], 1, d15.x + 40, d15.y + 30, 1.22],
        [20.7, 1, d15.x + 40, d15.y + 30, 1.22], [21.1, 0, list.x + list.w / 2, list.y + 160, 1],
        [K.pushCheck[0], 0.2, list.x + list.w / 2, list.y + 160, 1.1],
        [K.pushCheck[1], 1, list.x + list.w / 2, list.y + 180, 1.42], [K.handoff[0], 1, list.x + list.w / 2, list.y + 200, 1.42],
        [K.handoff[0] + 0.4, 0, list.x + list.w / 2, list.y + 200, 1],
      ];
      const [w, x, y, s] = Ease.keys(keys, t);
      return { w: t < K.pushAudit[0] ? 0 : w, x, y, s };
    },
    pointer(t) {
      const p = this.path;
      if (t < 9.0 || t > 22.0) {
        this.cursor.set(0, 0, 0, 0);
        this.ring.set(0, 0, 0);
        return;
      }
      let pos = { x: 1250, y: 860 }, press = 0, ring = null;
      for (let i = 0; i < p.length; i++) {
        const from = i === 0 ? 9.0 : p[i - 1].press + 0.2;
        const q = Ease.seg(t, from, p[i].at, Ease.inOutCubic);
        pos = { x: Ease.lerp(pos.x, p[i].x, q), y: Ease.lerp(pos.y, p[i].y, q) };
        press = Math.max(press, Ease.inOut(t, p[i].press - 0.08, p[i].press + 0.12, 0.06, 0.1));
        if (t >= p[i].press - 0.05 && t < p[i].press + 0.45) ring = { ...p[i], life: Ease.p(t, p[i].press - 0.05, p[i].press + 0.45) };
      }
      const pt = this.win.at(pos.x, pos.y);
      this.cursor.set(pt.x, pt.y, press, Ease.p(t, 9.0, 9.2) * (1 - Ease.p(t, 21.75, 22.0)));
      if (ring) {
        const c = this.win.at(ring.x, ring.y);
        this.ring.set(c.x, c.y, ring.life);
      } else this.ring.set(0, 0, 0);
    },
  });
})();
