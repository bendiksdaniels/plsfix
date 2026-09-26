//! track.rs tests: a picked track comes out exactly the video's length, faded at both ends,
//! untouched in the middle, and a short track leaves silence after its end.
use plsfix_video::core::sound::track::fit;

#[test]
fn a_long_track_is_cut_and_faded() {
    let src = vec![0.5f32; 48_000 * 70];
    let out = fit(&src, &src, 58.5);
    assert_eq!(out.len(), 48_000 * 58 + 24_000);
    assert_eq!(out.l[0], 0.0);
    assert_eq!(out.l[48_000 * 30], 0.5, "the middle is the track itself");
    assert!(out.l[out.len() - 48_000] < 0.5 && out.l[out.len() - 48_000] > 0.0, "fading out in the last 2,5 s");
    assert!(out.l[out.len() - 1].abs() < 1e-6);
}

#[test]
fn a_short_track_leaves_silence_after_its_end() {
    let src = vec![0.5f32; 48_000 * 10];
    let out = fit(&src, &src, 58.5);
    assert_eq!(out.len(), 48_000 * 58 + 24_000);
    assert_eq!(out.r[48_000 * 5], 0.5);
    assert_eq!(out.r[48_000 * 11], 0.0);
}
