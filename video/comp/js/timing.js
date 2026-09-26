// timing.js: when each scene runs, in seconds. Scenes overlap by 0.4 s so one can hand over
// to the next. Beats are 0.5 s (120 BPM): cuts, clicks and landings sit on them where they can.
window.T = {
  END: 87.5,
  BEAT: 0.5,
  intro: { from: 0, to: 6.4 },
  hero: { from: 6.0, to: 26.4 },
  tools: { from: 26.0, to: 60.4 },
  ppt: { from: 60.0, to: 80.4 },
  outro: { from: 79.8, to: 87.5 },
};
