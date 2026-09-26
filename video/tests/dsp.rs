//! dsp tests: the limiter holds its ceiling, noise is seeded and bounded, the swept filter stays
//! stable across its whole range, envelopes stay within 0..1.
use plsfix_video::core::dsp::dynamics::limit;
use plsfix_video::core::dsp::env::adsr;
use plsfix_video::core::dsp::filter::Svf;
use plsfix_video::core::dsp::osc::{Noise, Saw};

#[test]
fn the_limiter_holds_its_ceiling() {
    let mut l: Vec<f32> = (0..48_000).map(|i| (i as f32 * 0.05).sin() * if i % 5000 < 50 { 3.0 } else { 0.5 }).collect();
    let mut r = l.clone();
    limit(&mut l, &mut r, 0.794, 48_000.0, 0.002, 0.08);
    assert!(l.iter().chain(r.iter()).all(|x| x.abs() <= 0.7941));
}

#[test]
fn noise_is_seeded_and_bounded() {
    let (mut a, mut b) = (Noise::new(7), Noise::new(7));
    for _ in 0..10_000 {
        let x = a.next();
        assert_eq!(x, b.next());
        assert!((-1.0..1.0).contains(&x));
    }
}

#[test]
fn a_swept_filter_and_a_saw_stay_finite_and_bounded() {
    let mut saw = Saw::new(0.3);
    let mut f = Svf::default();
    for i in 0..96_000 {
        let cutoff = 20.0 * 1000f64.powf(i as f64 / 96_000.0);
        let x = saw.next(220.0 / 48_000.0);
        assert!(x.abs() <= 1.05);
        let y = f.tick(x as f64, cutoff, 0.8, 48_000.0).lp;
        assert!(y.is_finite() && y.abs() < 10.0);
    }
}

#[test]
fn envelopes_stay_in_range() {
    for i in 0..400 {
        let e = adsr(i as f64 * 0.01, 2.0, 0.45, 0.4, 0.8, 1.1);
        assert!((0.0..=1.0).contains(&e));
    }
    assert_eq!(adsr(-0.1, 2.0, 0.45, 0.4, 0.8, 1.1), 0.0);
    assert_eq!(adsr(3.2, 2.0, 0.45, 0.4, 0.8, 1.1), 0.0);
}

#[test]
fn the_limiter_glides_into_a_peak_instead_of_stepping() {
    // a steady 0.5 with one spike of 3.0: the gain must ramp down over the 96-sample look-ahead
    let mut l = vec![0.5f32; 4000];
    l[2000] = 3.0;
    let mut r = l.clone();
    limit(&mut l, &mut r, 0.794, 48_000.0, 0.002, 0.08);
    assert!(l[2000] <= 0.7941 && l[2000] > 0.79, "the peak lands on the ceiling: {}", l[2000]);
    let biggest_step = (1800..1999).map(|i| (l[i + 1] - l[i]).abs()).fold(0.0f32, f32::max);
    assert!(biggest_step < 0.01, "the gain steps by {biggest_step} in one sample");
    assert_eq!(l[1000], 0.5, "untouched well before the peak");
}
