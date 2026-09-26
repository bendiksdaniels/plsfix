// tools.js (26.0 - 60.4 s): "more than 100 tools". The hero window lands in the first slot of a
// 3 x 2 contact sheet, five more chapters fly in beside it with their chips; then each chapter
// takes the stage for 5.3 s (tools-beats.js): one real click, its result, a lean on it, and a
// caption tagged with the chapter's name, before it slides off for the next.
(() => {
  const CH = ToolsBeats.CHAPTERS;
  const TURN = 5.3;
  const START = 28.6;
  const K = { title: [26.45, 28.45], chips: 27.0, chipsOut: 28.45, fly: 26.2, sink: [28.6, 29.2] };
  const P = (o) => Object.freeze(o);
  const HERO = P({ x: 0, y: 0, s: 1, rx: 3, ry: -5 });
  const OFF_L = P({ x: -1800, y: 40, s: 0.85, rx: 0, ry: 30 });
  const OFF_R = P({ x: 1800, y: 40, s: 0.85, rx: 0, ry: -30 });
  const AIM = { x: 960, y: 500 };
  const turn = (i) => START + i * TURN;

  Stage.scene({
    id: 'tools', from: T.tools.from, to: T.tools.to,
    build(root) {
      this.title = UiType.headline(root, Copy.t('tools.title'), { left: '0px', right: '0px', top: '96px', textAlign: 'center', fontSize: '80px' });
      Stage.copy(this.title.el, 27.6, 'tools.title');
      Stage.cue(26.0, 'section', { name: 'lift' });
      Stage.voice('vo.tools.title', 26.3, 28.6);
      CH.forEach((ch, i) => Stage.voice(`vo.tools.${ch.key}`, turn(i) + 0.5, turn(i) + 4.9));
      this.items = CH.map((ch, i) => this.item(root, ch, i));
      CH.forEach((_, i) => i > 0 && Stage.cue(K.fly + i * 0.12, 'whoosh', { dur: 0.9, pan: i % 3 === 0 ? -0.6 : 0.6, gain: -14 }));
      CH.forEach((_, i) => Stage.cue(K.chips + i * 0.1 + 0.08, 'pop', { pitch: 72 + i * 2, gain: -11 }));
      CH.forEach((_, i) => Stage.cue(turn(i), 'whoosh', { dur: 0.7, gain: -6 }));
    },
    item(root, ch, i) {
      const win = UiWindow.create(root, Cam.HERO.cw, 'excel');
      const at = Cam.box(win, Cam.HERO.cx, Cam.HERO.cy);
      Cam.place(win, at);
      const beat = ToolsBeats.build(win, ch, turn(i));
      const cursor = UiPointer.cursor(win.content);
      const ring = UiPointer.ring(win.content);
      const name = Copy.t(`tools.${ch.key}`);
      const chip = UiType.roleChip(root, name, {});
      const slot = Cam.slot(i);
      const w = chip.el.getBoundingClientRect().width;
      Object.assign(chip.el.style, { left: `${Cam.HERO.cx + slot.x - w / 2}px`, top: `${Cam.HERO.cy + slot.y - win.h * slot.s / 2 - 26}px` });
      Stage.copy(chip.el, 27.9, `tools.${ch.key}`);
      const key = `tools.${ch.key}.1`;
      const cap = UiType.caption(root, Copy.t(key), null, name);
      Stage.copy(cap.el, turn(i) + 2.8, key);
      return { ch, i, win, at, beat, cursor, ring, chip, cap };
    },
    update(t) {
      this.title.update(t, K.title[0], K.title[1]);
      for (const it of this.items) {
        const pose = this.pose(it, t);
        const shown = this.shown(it, t);
        El.show(it.win.el, shown);
        it.win.el.style.zIndex = t >= turn(it.i) && t < turn(it.i) + TURN ? '3' : '1';
        if (shown > 0) {
          El.tf(it.win.el, pose);
          const prev = this.pose(it, t - 1 / 60);
          Cam.blur(it.win.el, `mb${it.i % 3}`, pose.x - prev.x, pose.y - prev.y);
          ToolsBeats.update(it.beat, it.win, t);
          this.pointer(it, t);
        }
        it.chip.update(t, K.chips + it.i * 0.1, K.chipsOut);
        it.cap.update(t, turn(it.i) + 0.55, turn(it.i) + 4.95);
      }
    },
    // visible while in the grid (the first fades in over the landing hero window) or on its turn
    shown(it, t) {
      const c = turn(it.i);
      if (t >= c - 0.05 && t < c + 5.6) return 1;
      if (it.i === 0) return t >= 26.3 && t < c ? Ease.p(t, 26.3, 26.45) : 0;
      return t >= K.fly && t < K.sink[1] ? 1 : 0;
    },
    pose(it, t) {
      const c = turn(it.i), slot = Cam.slot(it.i);
      const far = P({ ...slot, y: slot.y + 900, s: 0.3, rx: 30 });
      const gone = P({ ...slot, y: slot.y + 760, s: 0.28, rx: 20 });
      let pose;
      if (t < c - 0.05) {
        const fly = Ease.seg(t, K.fly + it.i * 0.12, K.fly + it.i * 0.12 + 0.9, Ease.outExpo);
        pose = it.i === 0 ? { ...slot } : Cam.mix(far, slot, fly);
        pose = it.i === 0 ? pose : Cam.mix(pose, gone, Ease.seg(t, K.sink[0], K.sink[1], Ease.inCubic));
        return pose;
      }
      pose = Cam.mix(it.i === 0 ? slot : OFF_R, HERO, Ease.seg(t, c, c + 0.7, Ease.outExpo));
      pose.ry += 2 * Ease.p(t, c + 0.7, c + 5);
      pose = this.lean(it, t, pose);
      return Cam.mix(pose, OFF_L, Ease.seg(t, c + 4.95, c + 5.55, Ease.inCubic));
    },
    lean(it, t, pose) {
      const c = turn(it.i);
      const w = Ease.seg(t, c + 1.8, c + 2.8, Ease.inOutCubic) * (1 - Ease.seg(t, c + 4.45, c + 4.95, Ease.inOutCubic));
      if (w <= 0) return pose;
      const sub = ToolsBeats.subject(it.beat);
      const s = sub.x > 1080 ? 1.32 : 1.16;
      const f = Cam.fit(it.win, it.at, s, Cam.focus(it.win, it.at, it.win.at(sub.x, sub.y), s, AIM));
      return Cam.mix(pose, { ...pose, s, x: f.x, y: f.y }, w);
    },
    pointer(it, t) {
      const c = turn(it.i), R = ToolsBeats.R;
      const to = ToolsBeats.target(it.beat);
      if (!to || t < c + 0.5 || t > c + 2.1) {
        it.cursor.set(0, 0, 0, 0);
        it.ring.set(0, 0, 0);
        return;
      }
      const q = Ease.seg(t, c + 0.5, c + R.press - 0.1, Ease.inOutCubic);
      const x = Ease.lerp(to.x + 220, to.x, q), y = Ease.lerp(to.y + 260, to.y, q);
      const press = Ease.inOut(t, c + R.press - 0.08, c + R.press + 0.12, 0.06, 0.1);
      const pt = it.win.at(x, y);
      it.cursor.set(pt.x, pt.y, press, Ease.p(t, c + 0.5, c + 0.7) * (1 - Ease.p(t, c + 1.9, c + 2.1)));
      const life = Ease.p(t, c + R.press - 0.05, c + R.press + 0.45);
      const r = it.win.at(to.x, to.y);
      it.ring.set(r.x, r.y, life > 0 && life < 1 ? life : 0);
    },
  });
})();
