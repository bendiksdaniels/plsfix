// clock.js: the master timeline. Scenes register with Stage.scene({id, from, to, build, update});
// __seek(t) shows the scenes whose [from, to) holds t and calls update(t) with absolute seconds.
// __ready runs every Stage.preload loader, builds the scenes, then waits for every <img> to
// decode and Montserrat to measure present, and resolves {duration, fps}.
// Stage.copy(el, t, id) registers on-screen text; __layoutCheck() seeks to each one's time and
// reports any whose text is not visible then, paints outside the title-safe area, or is
// covered by something else on screen. Stage.cue(t, kind, opts) registers a sound for the
// soundtrack at t (read back by __cues()), so the audio follows the same timing tables.
// Stage.voice(key, from, to) registers a voice-over line's slot (read back by __voice()): the
// narrator's words come from voice.<lang>.json, the time from the scene's own table.
window.Stage = (() => {
  const FPS = 60;
  const W = 1920;
  const H = 1080;
  const scenes = [];
  const loaders = [];
  const copies = [];
  const cues = [];
  const voices = [];
  let duration = 0;

  function scene(def) {
    scenes.push(def);
    duration = Math.max(duration, def.to);
  }

  function preload(fn) {
    loaders.push(fn);
  }

  function seek(t) {
    for (const s of scenes) {
      const on = t >= s.from && t < s.to;
      s.root.style.display = on ? '' : 'none';
      if (on) s.update(t);
    }
  }

  // kind: section | whoosh | click | tap | pop | impact | riser; opts: dur, pan, pitch, gain, name
  function voice(key, from, to) {
    voices.push({ key, from: Math.round(from * 1000) / 1000, to: Math.round(to * 1000) / 1000 });
  }

  function cue(t, kind, opts = {}) {
    cues.push({ t: Math.round(t * 1000) / 1000, kind, ...opts });
  }

  function copy(el, t, id) {
    el.dataset.copy = id;
    copies.push({ el, t, id });
  }

  const SAFE = 24;   // title-safe margin, px

  // the union of the boxes the element's text actually paints (its ink), in stage px
  function inkRect(el) {
    const range = document.createRange();
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    let r = null;
    while (walker.nextNode()) {
      const n = walker.currentNode;
      if (!n.textContent.trim()) continue;
      range.selectNodeContents(n);
      for (const b of range.getClientRects()) {
        if (b.width < 0.5 || b.height < 0.5) continue;
        r = r ? { left: Math.min(r.left, b.left), top: Math.min(r.top, b.top), right: Math.max(r.right, b.right), bottom: Math.max(r.bottom, b.bottom) }
          : { left: b.left, top: b.top, right: b.right, bottom: b.bottom };
      }
    }
    return r;
  }

  // the opacity the element really shows with (its own times every ancestor's)
  function shownOpacity(el) {
    let o = 1;
    for (let e = el; e && e !== document.body; e = e.parentElement) {
      const cs = getComputedStyle(e);
      if (cs.display === 'none' || cs.visibility === 'hidden') return 0;
      o *= Number(cs.opacity);
    }
    return o;
  }

  // what sits on top of the text, if anything: hit-tests the centre and four inner corners and
  // takes the topmost element that paints (containers marked data-pass paint nothing)
  function coveredBy(el, r) {
    const ix = Math.min(6, (r.right - r.left) / 4), iy = Math.min(6, (r.bottom - r.top) / 4);
    const pts = [[(r.left + r.right) / 2, (r.top + r.bottom) / 2], [r.left + ix, r.top + iy], [r.right - ix, r.top + iy], [r.left + ix, r.bottom - iy], [r.right - ix, r.bottom - iy]];
    for (const [x, y] of pts) {
      const top = document.elementsFromPoint(x, y).find((e) => !('pass' in e.dataset));
      if (top && top !== el && !el.contains(top)) return (top.className && String(top.className).split(' ')[0]) || top.tagName.toLowerCase();
    }
    return null;
  }

  function layoutCheck() {
    const problems = [];
    for (const c of copies) {
      seek(c.t);
      const at = (problem) => problems.push({ time: c.t, id: c.id, problem });
      const r = inkRect(c.el);
      if (!r || shownOpacity(c.el) < 0.6) {
        at('not visible at its own time');
        continue;
      }
      if (r.left < SAFE || r.top < SAFE || r.right > W - SAFE || r.bottom > H - SAFE) {
        at(`text leaves the title-safe area: ${Math.round(r.left)},${Math.round(r.top)} to ${Math.round(r.right)},${Math.round(r.bottom)}`);
      }
      const cover = coveredBy(c.el, r);
      if (cover) at(`text is covered by .${cover}`);
    }
    seek(0);
    return problems;
  }

  // A font that is missing falls back silently; measure it against the fallback instead.
  function fontPresent(family) {
    const c = document.createElement('canvas').getContext('2d');
    const probe = 'Projekti Uzdevumi Termiņi 0123';
    c.font = '700 40px monospace';
    const fallback = c.measureText(probe).width;
    c.font = `700 40px "${family}", monospace`;
    return Math.abs(c.measureText(probe).width - fallback) > 1;
  }

  async function ready() {
    for (const load of loaders) await load();
    const stage = document.getElementById('stage');
    for (const s of scenes) {
      s.root = document.createElement('div');
      s.root.className = 'scene';
      s.root.dataset.scene = s.id;
      s.root.dataset.pass = '';
      stage.appendChild(s.root);
      s.build(s.root);
    }
    await document.fonts.ready;
    if (!fontPresent('Montserrat')) throw new Error('font Montserrat is not installed');
    const imgs = [...document.images];
    await Promise.all(imgs.map((img) => img.decode().catch(() => {
      throw new Error(`image failed to load: ${img.getAttribute('src')}`);
    })));
    seek(0);
    return { duration, fps: FPS };
  }

  function boot() {
    window.__ready = ready();
    window.__seek = seek;
    window.__layoutCheck = layoutCheck;
    window.__cues = () => cues.slice().sort((a, b) => a.t - b.t);
    window.__voice = () => voices.slice().sort((a, b) => a.from - b.from);
  }

  return { scene, preload, seek, copy, cue, voice, boot, FPS, W, H };
})();
