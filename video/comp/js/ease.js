// ease.js: the motion vocabulary: progress over a time window, easing curves, a damped spring,
// fades in and out. Every function is pure (time in, number out); nothing reads a clock.
window.Ease = (() => {
  const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
  const lerp = (a, b, p) => a + (b - a) * p;
  // progress of t through [a, b], clamped to 0..1
  const p = (t, a, b) => (b <= a ? (t >= b ? 1 : 0) : clamp((t - a) / (b - a), 0, 1));

  const linear = (x) => x;
  const outQuad = (x) => 1 - (1 - x) * (1 - x);
  const inOutQuad = (x) => (x < 0.5 ? 2 * x * x : 1 - Math.pow(-2 * x + 2, 2) / 2);
  const outCubic = (x) => 1 - Math.pow(1 - x, 3);
  const inCubic = (x) => x * x * x;
  const inOutCubic = (x) => (x < 0.5 ? 4 * x * x * x : 1 - Math.pow(-2 * x + 2, 3) / 2);
  const outQuint = (x) => 1 - Math.pow(1 - x, 5);
  const outExpo = (x) => (x >= 1 ? 1 : 1 - Math.pow(2, -10 * x));
  const inExpo = (x) => (x <= 0 ? 0 : Math.pow(2, 10 * x - 10));
  const inOutExpo = (x) => (x <= 0 ? 0 : x >= 1 ? 1 : x < 0.5 ? Math.pow(2, 20 * x - 10) / 2 : (2 - Math.pow(2, -20 * x + 10)) / 2);
  const outBack = (x, s = 1.6) => 1 + (s + 1) * Math.pow(x - 1, 3) + s * Math.pow(x - 1, 2);

  // eased progress of t through [a, b]
  const seg = (t, a, b, curve = outCubic) => curve(p(t, a, b));

  // a damped spring settling from 0 to 1, `secs` after its start (overshoots when damp < 1)
  function spring(secs, freq = 2.2, damp = 0.55) {
    if (secs <= 0) return 0;
    const w = 2 * Math.PI * freq;
    const wd = w * Math.sqrt(1 - damp * damp);
    return 1 - Math.exp(-damp * w * secs) * (Math.cos(wd * secs) + (damp * w / wd) * Math.sin(wd * secs));
  }

  // 0 before `a`, eases to 1 over `inDur`, holds, eases back to 0 over `outDur` ending at `b`
  function inOut(t, a, b, inDur = 0.4, outDur = 0.4, curveIn = outCubic, curveOut = inCubic) {
    if (t < a || t >= b) return 0;
    if (t < a + inDur) return curveIn(p(t, a, a + inDur));
    if (t > b - outDur) return 1 - curveOut(p(t, b - outDur, b));
    return 1;
  }

  // keyframes [[t, v1, v2, ...], ...]: the values at time t, eased between neighbours
  function keys(list, t) {
    if (t <= list[0][0]) return list[0].slice(1);
    for (let i = 1; i < list.length; i++) {
      if (t <= list[i][0]) {
        const a = list[i - 1], b = list[i];
        const q = inOutCubic(p(t, a[0], b[0]));
        return a.slice(1).map((v, j) => lerp(v, b[j + 1], q));
      }
    }
    return list[list.length - 1].slice(1);
  }

  return { clamp, lerp, p, seg, spring, inOut, keys, linear, outQuad, inOutQuad, outCubic, inCubic, inOutCubic, outQuint, outExpo, inExpo, inOutExpo, outBack };
})();
