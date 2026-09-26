// ui-mark.js: the pls,fix mark (public/assets/icon.svg: a navy rounded square, a mint comma in
// Georgia bold) drawn inline with the comma's head position known, so the intro's mint drop can
// grow out of it; plus the "pls,fix" wordmark with the same mint comma.
window.UiMark = (() => {
  const NS = 'http://www.w3.org/2000/svg';

  // size = the square's side in px; the mark is public/assets/icon.svg drawn inline, so its
  // comma sits exactly where the add-in's own icon has it. `head` = the comma's head, in px.
  function create(parent, size) {
    const el = El.make('div', 'mark', parent, { width: `${size}px`, height: `${size}px` });
    El.make('div', 'mark-rim', el, { borderRadius: `${size * 14 / 64}px` });
    const svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', '0 0 64 64');
    el.appendChild(svg);
    const sq = document.createElementNS(NS, 'rect');
    Object.entries({ width: 64, height: 64, rx: 14, fill: '#14213D' }).forEach(([k, v]) => sq.setAttribute(k, v));
    svg.appendChild(sq);
    const comma = document.createElementNS(NS, 'text');
    Object.entries({ x: 32, y: 44, 'font-family': 'Georgia, serif', 'font-size': 66, 'font-weight': 700, 'text-anchor': 'middle', fill: '#2EC4B6' }).forEach(([k, v]) => comma.setAttribute(k, v));
    comma.textContent = ',';
    svg.appendChild(comma);
    const b = comma.getBBox();
    const head = { x: (b.x + b.width * 0.55) * size / 64, y: (b.y + b.height * 0.32) * size / 64 };
    return { el, comma, size, head };
  }

  function wordmark(parent, px) {
    const el = El.make('div', 'wordmark', parent, { fontSize: `${px}px` });
    const inner = El.make('span', 'wordmark-in', el);
    inner.innerHTML = '<span class="wm-a">pls</span><span class="wm-comma">,</span>fix';
    return { el, inner, first: inner.firstChild };
  }

  // the mark and the wordmark side by side, centred on (cx, cy); `shift` = how far right the
  // mark sits when it is alone in the middle (before the wordmark comes out)
  function lockup(parent, size, wordPx, gap, cx, cy) {
    const mark = create(parent, size);
    const word = wordmark(parent, wordPx);
    mark.el.style.zIndex = '2'; // the wordmark slides out from behind the mark
    const b = word.el.getBoundingClientRect();
    const width = size + gap + b.width;
    const x = cx - width / 2, y = cy - size / 2;
    word.el.style.left = `${x + size + gap}px`;
    word.el.style.top = `${cy - b.height / 2}px`;
    return { mark, word, x, y, shift: width / 2 - size / 2 };
  }

  return { create, wordmark, lockup };
})();
