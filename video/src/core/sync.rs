//! sync.rs: how far one recording is shifted against another: the lag (in samples) at which
//! their cross-correlation peaks. The audit uses it to prove the file's soundtrack sits exactly
//! where the master was written (the master itself is timed from the picture's own cues).

/// The lag of `b` against `a` within +-`max_lag` samples: positive = `b` comes later.
pub fn lag(a: &[f32], b: &[f32], max_lag: usize) -> isize {
    let mut best = (0isize, f64::MIN);
    let m = max_lag as isize;
    for d in -m..=m {
        let mut sum = 0.0f64;
        for (i, &x) in a.iter().enumerate() {
            let j = i as isize + d;
            if j >= 0 && (j as usize) < b.len() {
                sum += x as f64 * b[j as usize] as f64;
            }
        }
        if sum > best.1 {
            best = (d, sum);
        }
    }
    best.0
}
