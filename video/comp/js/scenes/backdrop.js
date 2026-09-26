// backdrop.js: the ground under every scene for the whole video: a deep navy radial field, a
// drifting dot grid, a mint glow that follows each scene's focus, fine grain against banding.
Stage.scene({
  id: 'backdrop', from: 0, to: T.END,
  build(root) {
    El.make('div', 'bd-ground', root);
    this.grid = El.make('div', 'bd-grid', root);
    this.glow = El.make('div', 'bd-glow', root);
    El.make('div', 'bd-grain', root, { backgroundImage: `url(${grain()})` });
    El.make('div', 'bd-vignette', root);
  },
  update(t) {
    this.grid.style.transform = `translate(${-(t * 9) % 32}px, ${-(t * 5) % 32}px)`;
    const [x, y] = Ease.keys([[0, 960, 420], [6, 900, 520], [16, 1000, 560], [26, 960, 560], [33, 820, 540], [44, 1100, 540], [56, 860, 560], [64, 700, 540], [72, 1250, 540], [80, 960, 500], [87.5, 960, 470]], t);
    El.tf(this.glow, { x: x - 700, y: y - 700 });
  },
});

// a fixed, seeded noise tile: the same grain in every render and every worker
function grain() {
  const c = document.createElement('canvas');
  c.width = c.height = 256;
  const g = c.getContext('2d');
  const d = g.createImageData(256, 256);
  let s = 20260926;
  const rnd = () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 4294967296; };
  for (let i = 0; i < d.data.length; i += 4) {
    const v = 96 + Math.round(rnd() * 64);
    d.data[i] = d.data[i + 1] = d.data[i + 2] = v;
    d.data[i + 3] = 255;
  }
  g.putImageData(d, 0, 0);
  return c.toDataURL('image/png');
}
