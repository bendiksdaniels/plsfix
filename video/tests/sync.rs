//! sync.rs tests: the lag of a shifted copy is found exactly, both ways.
use plsfix_video::core::sync::lag;

fn signal() -> Vec<f32> {
    (0..4000).map(|i| ((i as f32 * 0.37).sin() * (i as f32 * 0.011).cos()) + if i == 1500 { 3.0 } else { 0.0 }).collect()
}

#[test]
fn a_shifted_copy_is_found() {
    let a = signal();
    let mut later = vec![0.0; 25];
    later.extend_from_slice(&a[..a.len() - 25]);
    assert_eq!(lag(&a, &later, 60), 25);
    assert_eq!(lag(&later, &a, 60), -25);
    assert_eq!(lag(&a, &a, 60), 0);
}
