// intro.js (0 - 6.4 s): the name's joke. A mail from the MD drops in: "pls fix". The camera
// leans in, a mint comma falls between the words ("pls,fix"), the card melts away and the text
// flies up into the wordmark as the navy mark pops in beside it; three words, the claim; then
// a mint drop grows out of the mark's comma until it fills the frame (the hero scene lifts it).
(() => {
  const MARK = 170;          // the mark's side, px
  const WORD_PX = 150;
  const GAP = 44;
  const CY = 400;            // the lockup's vertical centre
  const CARD = { x: 580, y: 330 };
  const K = {
    card: [0.3, 1.0], lean: [0.95, 1.6], drop: [1.55, 1.95], swap: 2.25, fly: [2.25, 2.9], mark: 2.55,
    words: [2.85, 4.4], wordGap: 0.22, claim: [4.85, 5.9], dive: [5.95, 6.4],
  };
  const LEAN = 1.32;         // the card's scale once the camera leans in

  Stage.scene({
    id: 'intro', from: T.intro.from, to: T.intro.to,
    build(root) {
      this.group = El.make('div', 'layer', root);
      this.group.dataset.pass = '';
      this.mail = UiMail.create(this.group);
      Object.assign(this.mail.el.style, { left: `${CARD.x}px`, top: `${CARD.y}px` });
      const lk = UiMark.lockup(this.group, MARK, WORD_PX, GAP, 960, CY);
      Object.assign(this, { mark: lk.mark, word: lk.word, markX: lk.x, markY: lk.y });
      this.word.el.style.transformOrigin = '0 0';
      this.fly = this.flight();
      this.drop = El.make('div', 'fill', this.group, { background: 'var(--mint)', borderRadius: '50%', inset: 'auto', width: '100px', height: '100px' });
      this.aim = { x: this.markX + this.mark.head.x, y: this.markY + this.mark.head.y };
      Object.assign(this.drop.style, { left: `${this.aim.x - 50}px`, top: `${this.aim.y - 50}px` });
      const line = { left: '0px', right: '0px', top: `${CY + 150}px`, textAlign: 'center' };
      this.words = UiType.headline(this.group, Copy.t('intro.words'), { ...line, fontSize: '64px', color: 'var(--ink-2)', wordSpacing: '0.45em' }, K.wordGap);
      this.claim = UiType.headline(this.group, Copy.t('intro.claim'), { ...line, fontSize: '84px', top: `${CY + 138}px` });
      this.cues();
      Stage.copy(this.words.el, 3.9, 'intro.words');
      Stage.copy(this.claim.el, 5.5, 'intro.claim');
    },
    cues() {
      Stage.cue(0, 'section', { name: 'intro' });
      Stage.cue(K.card[0], 'whoosh', { dur: 0.6, pan: -0.2, gain: -12 });
      Stage.cue(K.card[0] + 0.45, 'pop', { pitch: 76, gain: -6 });
      Stage.cue(K.drop[1] - 0.05, 'tap', { gain: -2 });
      Stage.cue(K.drop[1] - 0.05, 'pop', { pitch: 84, gain: -8 });
      Stage.cue(K.fly[0], 'whoosh', { dur: 0.6, pan: 0.3, gain: -10 });
      Stage.cue(K.mark + 0.05, 'pop', { pitch: 72, gain: -3 });
      [0, 1, 2].forEach((i) => Stage.cue(K.words[0] + i * K.wordGap + 0.05, 'click', { gain: -8 }));
      Stage.cue(K.claim[0] + 0.05, 'pop', { pitch: 84, gain: -8 });
      Stage.cue(K.dive[1] - 1.0, 'riser', { dur: 1.0, gain: -4 });
      Stage.cue(K.dive[1], 'impact', { gain: 0 });
    },
    // where the mail's "pls" sits at the swap (the card leaned in about its centre), as the
    // translation and scale that put the wordmark's own "pls" exactly there
    flight() {
      const card = this.mail.el.getBoundingClientRect();
      const c = { x: card.left + card.width / 2, y: card.top + card.height / 2 };
      const a = this.mail.a.getBoundingClientRect();
      const b = this.word.first.getBoundingClientRect();
      const w = this.word.el.getBoundingClientRect();
      const s = (a.height * LEAN) / b.height;
      const ax = c.x + LEAN * (a.left - c.x), ay = c.y + LEAN * (a.top - c.y);
      return { s, x: ax - w.left - s * (b.left - w.left), y: ay - w.top - s * (b.top - w.top) };
    },
    update(t) {
      this.card(t);
      const fly = Ease.seg(t, K.fly[0], K.fly[1], Ease.outExpo);
      El.show(this.word.el, t >= K.swap ? 1 : 0);
      El.tf(this.word.el, { x: this.fly.x * (1 - fly), y: this.fly.y * (1 - fly), s: Ease.lerp(this.fly.s, 1, fly) });
      const ink = Ease.seg(t, K.swap, K.swap + 0.4);
      this.word.el.style.color = `rgb(${Ease.lerp(20, 244, ink)}, ${Ease.lerp(33, 244, ink)}, ${Ease.lerp(61, 244, ink)})`;
      const pop = Ease.clamp(Ease.spring(t - K.mark, 2.4, 0.5), 0, 1.2);
      El.show(this.mark.el, t >= K.mark ? 1 : 0);
      El.tf(this.mark.el, { x: this.markX, y: this.markY + (1 - Math.min(1, pop)) * 30, s: pop });
      this.mark.el.style.transformOrigin = '50% 50%';
      this.words.update(t, K.words[0], K.words[1]);
      this.claim.update(t, K.claim[0], K.claim[1] + 3);
      const d = Ease.seg(t, K.dive[0], K.dive[1], Ease.inExpo);
      El.show(this.drop, t >= K.dive[0] ? 1 : 0);
      El.tf(this.drop, { s: 0.01 + 46 * d });
      El.tf(this.group, { x: (960 - this.aim.x) * 0.2 * d, y: (540 - this.aim.y) * 0.2 * d, s: 1 + 0.25 * d });
    },
    // the card: drops in, the camera leans in, the comma falls into the gap, the card melts
    card(t) {
      const m = this.mail;
      const drop = Ease.clamp(Ease.spring(t - K.card[0], 2.0, 0.62), 0, 1.1);
      const lean = Ease.seg(t, K.lean[0], K.lean[1], Ease.inOutCubic);
      El.show(m.el, t >= K.card[0] && t < K.swap + 0.4 ? 1 : 0);
      El.tf(m.el, { y: (1 - drop) * -520, s: 1 + (LEAN - 1) * lean });
      // at the swap the words become the wordmark (drawn on top); the card around them melts
      El.show(m.text, t < K.swap ? 1 : 0);
      // the header and icon go first, so the wordmark never flies up through the subject line
      El.show(m.box, 1 - Ease.seg(t, K.swap, K.swap + 0.35, Ease.outQuad));
      [m.icon, m.head].forEach((e) => El.show(e, 1 - Ease.seg(t, K.swap - 0.05, K.swap + 0.12, Ease.outQuad)));
      const fall = Ease.seg(t, K.drop[0], K.drop[1], Ease.inCubic);
      const squash = Ease.seg(t, K.drop[1] - 0.02, K.drop[1] + 0.12);
      m.slot.style.width = `${Ease.lerp(m.space, m.commaW, Ease.outCubic(squash))}px`;
      El.show(m.comma, t >= K.drop[0] ? 1 : 0);
      El.tf(m.comma, { y: (1 - fall) * -150 });
    },
  });
})();
