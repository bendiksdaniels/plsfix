// shots.js: the capture manifest (video/build/shots/shots.json) and the on-screen text
// (copy.lv.json), loaded before any scene builds; shots come only from the demo build. Shots.rect(name, key) gives a capture's
// rectangle in CSS px; Copy.t(key) gives a caption. A missing name fails loudly.
window.Shots = (() => {
  const BASE = '../build/shots/';
  let manifest = null;

  async function load() {
    const r = await fetch(BASE + 'shots.json');
    if (!r.ok) throw new Error('shots.json missing: run plsfix-video capture first');
    manifest = await r.json();
    if (manifest.source !== 'npm run demo:build') throw new Error(`shots.json source is ${manifest.source}, want npm run demo:build`);
  }

  function shot(name) {
    const s = manifest && manifest.shots[name];
    if (!s) throw new Error(`no shot ${name} in shots.json`);
    return s;
  }

  const src = (name) => BASE + shot(name).file;

  // one rectangle {x, y, w, h}; for a list key, `i` picks the entry
  function rect(name, key, i) {
    const r = shot(name).rects[key];
    if (!r) throw new Error(`no rect ${key} in shot ${name}`);
    const v = Array.isArray(r[0]) ? r[i || 0] : r;
    if (!v) throw new Error(`rect ${key}[${i}] missing in shot ${name}`);
    return { x: v[0], y: v[1], w: v[2], h: v[3] };
  }

  function rects(name, key) {
    const r = shot(name).rects[key];
    if (!r || !Array.isArray(r[0])) throw new Error(`no rect list ${key} in shot ${name}`);
    return r.map((v) => ({ x: v[0], y: v[1], w: v[2], h: v[3] }));
  }

  Stage.preload(load);
  // the rect keys a shot has (a list of hits whose count the capture decided)
  const keys = (name) => Object.keys(shot(name).rects);

  return { shot, src, rect, rects, keys };
})();

window.Copy = (() => {
  let text = null;
  async function load() {
    const r = await fetch('copy.lv.json');
    if (!r.ok) throw new Error('copy.lv.json missing');
    text = await r.json();
  }
  function t(key) {
    if (!text || !(key in text)) throw new Error(`no copy for ${key}`);
    return text[key];
  }
  Stage.preload(load);
  return { t };
})();
