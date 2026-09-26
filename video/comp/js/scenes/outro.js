// outro.js (79.8 - 87.5 s): the end card. The mark pops in alone in the middle as the
// PowerPoint act leaves, slides left as "pls,fix" comes out from behind it, the intro's claim
// returns, then "free, open source" and the GitHub address, held about 4 s to read; the frame
// leans in slowly until the curtain closes over the last 0.3 s.
(() => {
  const K = { pop: 80.15, slide: [80.7, 81.3], claim: [81.1, 90], open: [81.8, 82.35], url: [82.3, 82.9] };
  const MARK = 150;
  const WORD_PX = 128;

  Stage.scene({
    id: 'outro', from: T.outro.from, to: T.outro.to,
    build(root) {
      this.group = El.make('div', 'layer', root);
      this.group.dataset.pass = '';
      this.group.style.transformOrigin = '960px 540px';
      const lk = UiMark.lockup(this.group, MARK, WORD_PX, 40, 960, 360);
      Object.assign(this, { mark: lk.mark, word: lk.word, x: lk.x, y: lk.y, shift: lk.shift });
      this.mark.el.style.transformOrigin = '50% 50%';
      this.claim = UiType.headline(this.group, Copy.t('intro.claim'), { left: '0px', right: '0px', top: '488px', textAlign: 'center', fontSize: '66px' });
      this.open = El.make('div', 'end-open', this.group, { top: '604px' });
      this.open.textContent = Copy.t('outro.open');
      this.url = El.make('div', 'end-url', this.group, { top: '684px' });
      El.make('span', 'end-gh', this.url).textContent = Copy.t('outro.github');
      El.make('span', null, this.url).textContent = Copy.t('outro.url');
      Stage.copy(this.claim.el, 85.0, 'outro.claim');
      Stage.copy(this.open, 85.0, 'outro.open');
      Stage.copy(this.url, 85.0, 'outro.url');
      Stage.cue(80.0, 'section', { name: 'outro' });
      Stage.voice('vo.outro.claim', 80.8, 83.3);
      Stage.voice('vo.outro.open', 83.5, 86.9);
      Stage.cue(K.pop + 0.05, 'pop', { pitch: 72, gain: -3 });
      Stage.cue(K.slide[0], 'whoosh', { dur: 0.55, pan: 0.35, gain: -10 });
      Stage.cue(K.slide[0] + 0.05, 'impact', { gain: -9 });
      Stage.cue(K.open[0] + 0.05, 'pop', { pitch: 86, gain: -11 });
      Stage.cue(K.url[0] + 0.05, 'pop', { pitch: 91, gain: -10 });
    },
    update(t) {
      const pop = Ease.clamp(Ease.spring(t - K.pop, 2.4, 0.5), 0, 1.2);
      const slide = Ease.seg(t, K.slide[0], K.slide[1], Ease.outExpo);
      El.show(this.mark.el, t >= K.pop ? 1 : 0);
      El.tf(this.mark.el, { x: this.x + this.shift * (1 - slide), y: this.y, s: pop });
      El.tf(this.word.inner, { x: -(1 - slide) * 110 * WORD_PX / 30 });
      El.show(this.word.el, t >= K.slide[0] ? 1 : 0);
      this.claim.update(t, K.claim[0], K.claim[1]);
      const o = Ease.seg(t, K.open[0], K.open[1], Ease.outExpo);
      El.show(this.open, t >= K.open[0] ? o : 0);
      El.tf(this.open, { y: (1 - o) * 20 });
      const u = Ease.seg(t, K.url[0], K.url[1], Ease.outExpo);
      El.show(this.url, t >= K.url[0] ? u : 0);
      this.url.style.transform = `translate(-50%, ${(1 - u) * 24}px)`;
      El.tf(this.group, { s: 1 + 0.045 * Ease.p(t, K.pop, T.END) });
    },
  });
})();
