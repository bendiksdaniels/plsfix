//! dynamics.rs: the master's peak limiter (look-ahead, so it turns down before a peak instead of
//! clipping it) and gain helpers. Invariant: after `limit`, no sample exceeds the ceiling.

use std::collections::VecDeque;

pub fn db_to_gain(db: f64) -> f64 {
    10f64.powf(db / 20.0)
}

/// Scales both channels in place.
pub fn gain(l: &mut [f32], r: &mut [f32], g: f64) {
    let g = g as f32;
    for x in l.iter_mut().chain(r.iter_mut()) {
        *x *= g;
    }
}

/// Look-ahead limiting to `ceiling` (linear, e.g. -2 dBFS = 0.794). The gain glides down over
/// the `look_s` before each peak (a moving average of the held minimum, so it reaches the peak's
/// need exactly at the peak and never steps, which would click) and recovers over `release_s`.
pub fn limit(l: &mut [f32], r: &mut [f32], ceiling: f32, sr: f64, look_s: f64, release_s: f64) {
    let n = l.len().min(r.len());
    let look = ((look_s * sr) as usize).max(1);
    let need: Vec<f32> = (0..n).map(|i| {
        let p = l[i].abs().max(r[i].abs());
        if p > ceiling { ceiling / p } else { 1.0 }
    }).collect();
    let hold = min_ahead(&need, look);
    let rel = (-1.0 / (release_s * sr)).exp() as f32;
    let (mut g, mut sum) = (1.0f32, (look + 1) as f64);
    for i in 0..n {
        // the window [i - look, i] of held minimums; before the start it counts as 1 (no limiting)
        sum += hold[i] as f64 - if i > look { hold[i - look - 1] as f64 } else { 1.0 };
        let attack = (sum / (look + 1) as f64) as f32;
        g = if attack < g { attack } else { attack + (g - attack) * rel };
        l[i] = (l[i] * g).clamp(-ceiling, ceiling);
        r[i] = (r[i] * g).clamp(-ceiling, ceiling);
    }
}

/// The smallest value in [i, i + look] for every i (a sliding-window minimum, O(n)).
fn min_ahead(v: &[f32], look: usize) -> Vec<f32> {
    let mut window: VecDeque<usize> = VecDeque::new();
    let mut out = vec![1.0f32; v.len()];
    for j in (0..v.len()).rev() {
        while window.back().is_some_and(|&k| v[k] >= v[j]) {
            window.pop_back();
        }
        window.push_back(j);
        while window.front().is_some_and(|&k| k > j + look) {
            window.pop_front();
        }
        out[j] = v[*window.front().unwrap_or(&j)];
    }
    out
}
