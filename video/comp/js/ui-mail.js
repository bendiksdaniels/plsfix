// ui-mail.js: the intro's mail notification ("pls fix" from the MD) as separate pieces: the
// card's box and header fade on their own, the two words and the comma stay, so the comma can
// fall between the words and the text can become the wordmark. Returns the parts; no motion.
window.UiMail = (() => {
  const ENVELOPE = '<svg viewBox="0 0 64 64" width="64" height="64" aria-hidden="true"><rect x="13" y="19" width="38" height="27" rx="5" fill="#fff"/><path d="M15 22 L32 35 L49 22" fill="none" stroke="#2F6BD8" stroke-width="3.2" stroke-linejoin="round"/></svg>';

  function create(parent) {
    const el = El.make('div', 'mail', parent);
    el.dataset.pass = '';
    // the card's face and shadow live on the box, so the whole card melts when the box fades
    const box = El.make('div', 'fill', el, { borderRadius: '26px', background: 'rgba(247, 248, 250, 0.97)', boxShadow: '0 40px 90px rgba(0, 0, 0, 0.5)' });
    const icon = El.make('div', 'mail-icon', el);
    icon.innerHTML = ENVELOPE;
    const body = El.make('div', 'mail-body', el);
    const head = El.make('div', null, body);
    const top = El.make('div', 'mail-top', head);
    El.make('span', null, top).textContent = Copy.t('intro.mail.from');
    El.make('span', 'mail-time', top).textContent = Copy.t('intro.mail.time');
    El.make('div', 'mail-subject', head).textContent = Copy.t('intro.mail.subject');
    // "pls" [slot] "fix": the slot starts as the space between the words and narrows to the
    // comma's own width as the comma lands in it, baseline-aligned because it is inline
    const text = El.make('div', 'mail-text', body);
    const [w1, w2] = Copy.t('intro.mail.body').split(' ');
    const a = El.make('span', 'mail-word', text);
    a.textContent = w1;
    const slot = El.make('span', 'mail-slot', text);
    const comma = El.make('span', 'mail-comma', slot);
    comma.textContent = ',';
    const b = El.make('span', 'mail-word', text);
    b.textContent = w2;
    const space = measure(text, '\u00a0', 'var(--font)');
    const commaW = measure(text, ',', 'var(--serif)');
    return { el, box, icon, head, text, a, slot, comma, b, space, commaW };
  }

  // the advance width of `ch` in the text's own size and weight
  function measure(text, ch, family) {
    const probe = El.make('span', null, text, { position: 'absolute', visibility: 'hidden', fontFamily: family });
    probe.textContent = ch;
    const w = probe.getBoundingClientRect().width;
    probe.remove();
    return w;
  }

  return { create };
})();
