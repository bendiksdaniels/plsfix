// ui-type.js: kinetic type. headline() reveals words one by one from behind a mask and sends
// them up and out; caption() is the pill at the bottom of the frame; roleChip() is the gold
// label over a window. *word* marks an accent (gold) or, in a caption, an English UI name.
window.UiType = (() => {
  // "a *b c* d\ne" -> lines of [{text, accent}]; an accent may span several words
  function parse(text) {
    let on = false;
    return text.split('\n').map((line) => line.split(' ').filter(Boolean).map((w) => {
      const opens = w.startsWith('*'), closes = w.slice(1).includes('*');
      const accent = on || opens;
      if (opens) on = true;
      if (closes) on = false;
      return { text: w.replace(/\*/g, ''), accent };
    }));
  }

  function words(el, text) {
    const spans = [];
    parse(text).forEach((line, li) => {
      if (li > 0) El.make('br', null, el);
      line.forEach((w, wi) => {
        const mask = El.make('span', 'wmask', el);
        const inner = El.make('span', `w${w.accent ? ' acc' : ''}`, mask);
        inner.textContent = w.text;
        spans.push(inner);
        if (wi < line.length - 1) el.appendChild(document.createTextNode(' '));
      });
    });
    return spans;
  }

  // each word rises in over 0.55 s (stagger `gap`), and leaves upward over 0.35 s from tOut
  // (stagger `outGap`)
  function reveal(spans, t, tIn, tOut, gap, outGap = gap * 0.5) {
    spans.forEach((s, i) => {
      const pin = Ease.seg(t, tIn + i * gap, tIn + i * gap + 0.55, Ease.outExpo);
      const pout = Ease.seg(t, tOut + i * outGap, tOut + i * outGap + 0.35, Ease.inCubic);
      El.tf(s, { y: (1 - pin) * 105 - pout * 105 });
      El.show(s, t < tIn ? 0 : 1);
    });
  }

  function headline(parent, text, css, gap = 0.07, outGap = 0.03) {
    const el = El.make('div', 'headline', parent, css);
    const spans = words(el, text);
    function update(t, tIn, tOut) {
      El.show(el, t >= tIn && t < tOut + 0.6 ? 1 : 0);
      reveal(spans, t, tIn, tOut, gap, outGap);
    }
    return { el, update };
  }

  // `tag` (a role name) replaces the gold square with a gold label
  function caption(parent, text, css, tag) {
    const el = El.make('div', 'pill', parent, css);
    if (tag) El.make('span', 'pill-tag', el).textContent = tag;
    else El.make('span', 'pill-mark', el);
    const body = El.make('span', 'pill-text', el);
    const spans = words(body, text);
    function update(t, tIn, tOut) {
      const p = Ease.inOut(t, tIn, tOut, 0.45, 0.35, Ease.outExpo, Ease.inCubic);
      El.show(el, p);
      el.style.transform = `translate(-50%, ${(1 - p) * 26}px)`;
      reveal(spans, t, tIn + 0.08, tOut + 5, 0.035);
    }
    return { el, update };
  }

  function roleChip(parent, text, css) {
    const el = El.make('div', 'role-chip', parent, css);
    el.textContent = text;
    function update(t, tIn, tOut) {
      const p = Ease.inOut(t, tIn, tOut, 0.4, 0.3, Ease.outBack, Ease.inCubic);
      El.show(el, Math.min(1, p));
      El.tf(el, { y: (1 - p) * 14, s: 0.9 + 0.1 * p });
    }
    return { el, update };
  }

  return { headline, caption, roleChip, parse };
})();
