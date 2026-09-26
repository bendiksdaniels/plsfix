// curtain.js: the top layer: the opening fade from near-black and the closing fade to it.
Stage.scene({
  id: 'curtain', from: 0, to: T.END,
  build(root) {
    this.el = El.make('div', 'curtain', root);
  },
  update(t) {
    const open = 1 - Ease.seg(t, 0, 0.45, Ease.outQuad);
    const close = Ease.seg(t, T.END - 0.35, T.END - 0.03, Ease.inQuad);
    El.show(this.el, Math.max(open, close));
  },
});
