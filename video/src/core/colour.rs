//! colour.rs: the colour gate's arithmetic: the mean colour of a small patch of an RGB frame,
//! and the probe points where the browser's frame and the decoded video must agree.
//! Invariant: a probe sits on a flat area, so encoding noise stays well under the tolerance.

/// Where the audit compares the browser's frame with the decoded video: time (a whole frame at
/// 60 fps), x, y, and what is there.
pub const PROBES: &[(f64, u32, u32, &str)] = &[
    (2.0, 160, 160, "the navy ground (intro)"),
    (6.4, 960, 540, "the full-frame mint"),
    (47.0, 80, 1030, "the ground under the toolbox"),
    (84.0, 200, 950, "the ground under the end card"),
];

/// Largest allowed difference per channel between the two patch means.
pub const TOLERANCE: f64 = 5.0;

/// The mean RGB of the (2r+1)^2 patch centred on (x, y) of a `w`-wide RGB frame.
pub fn patch_mean(rgb: &[u8], w: u32, x: u32, y: u32, r: u32) -> Option<[f64; 3]> {
    let mut sum = [0.0; 3];
    let mut n = 0.0;
    for py in y.saturating_sub(r)..=y + r {
        for px in x.saturating_sub(r)..=x + r {
            let i = ((py * w + px) * 3) as usize;
            let p = rgb.get(i..i + 3)?;
            for c in 0..3 {
                sum[c] += p[c] as f64;
            }
            n += 1.0;
        }
    }
    Some([sum[0] / n, sum[1] / n, sum[2] / n])
}

/// How many pixels of two same-size RGB frames differ by more than `level` in some channel.
/// Chrome's own gradient dithering moves a pixel by 1; a leaked state moves thousands by a lot.
pub fn pixels_differing(a: &[u8], b: &[u8], level: u8) -> usize {
    if a.len() != b.len() {
        return a.len().max(b.len()) / 3;
    }
    a.chunks_exact(3).zip(b.chunks_exact(3)).filter(|(p, q)| p.iter().zip(q.iter()).any(|(x, y)| x.abs_diff(*y) > level)).count()
}

/// The determinism gate: a frame may differ from its twin by renderer noise only. Chrome dithers
/// gradients by 1 and, with 30 screenshots of 2880x1800 in memory, resamples a large downscaled
/// shot a little differently from one page's history to another's (measured 26.09.2026: at most
/// 18 levels, on text edges across a whole window). A leaked state moves, shows or hides
/// something, and those pixels change by far more.
pub const NOISE_LEVEL: u8 = 32;
/// More pixels than this beyond the noise level is a finding (about a 14 x 14 px patch).
pub const MAX_CONTENT_PIXELS: usize = 200;

/// The largest per-channel difference between two colours.
pub fn max_diff(a: [f64; 3], b: [f64; 3]) -> f64 {
    (0..3).map(|c| (a[c] - b[c]).abs()).fold(0.0, f64::max)
}
