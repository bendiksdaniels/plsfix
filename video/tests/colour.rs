//! colour.rs tests: the patch mean and the difference the colour gate compares.
use plsfix_video::core::colour::{max_diff, patch_mean, PROBES};

#[test]
fn patch_mean_averages_the_square() {
    // a 3x3 frame: every pixel (10, 20, 30) except the centre (40, 50, 60)
    let mut rgb = vec![0u8; 27];
    for i in 0..9 {
        let v: [u8; 3] = if i == 4 { [40, 50, 60] } else { [10, 20, 30] };
        rgb[i * 3..i * 3 + 3].copy_from_slice(&v);
    }
    let m = patch_mean(&rgb, 3, 1, 1, 1).unwrap();
    assert!((m[0] - 13.333).abs() < 0.01 && (m[2] - 33.333).abs() < 0.01);
    assert!(patch_mean(&rgb, 3, 2, 2, 1).is_none(), "a patch past the frame is refused");
}

#[test]
fn max_diff_is_the_worst_channel() {
    assert_eq!(max_diff([10.0, 20.0, 30.0], [12.0, 14.0, 31.0]), 6.0);
}

#[test]
fn probes_sit_on_whole_frames_inside_the_stage() {
    for &(t, x, y, _) in PROBES {
        assert!(((t * 60.0) - (t * 60.0).round()).abs() < 1e-9, "{t} is not a frame time");
        assert!(x >= 4 && x < 1916 && y >= 4 && y < 1076);
    }
}

#[test]
fn dithering_is_not_a_difference_but_a_moved_block_is() {
    use plsfix_video::core::colour::pixels_differing;
    let a = vec![100u8; 30];
    let mut dither = a.clone();
    dither[4] = 101;
    assert_eq!(pixels_differing(&a, &dither, 8), 0);
    let mut moved = a.clone();
    moved[0..9].copy_from_slice(&[200; 9]);
    assert_eq!(pixels_differing(&a, &moved, 8), 3);
    assert_eq!(pixels_differing(&a, &a[..27], 8), 10, "a size mismatch counts every pixel");
}

#[test]
fn determinism_noise_is_not_content() {
    use plsfix_video::core::colour::{pixels_differing, MAX_CONTENT_PIXELS, NOISE_LEVEL};
    // resampling noise: every pixel of a window off by up to 18 levels
    let a = vec![120u8; 3 * 10_000];
    let noisy: Vec<u8> = a.iter().enumerate().map(|(i, v)| if i % 2 == 0 { v + 18 } else { v - 18 }).collect();
    assert_eq!(pixels_differing(&a, &noisy, NOISE_LEVEL), 0);
    // a leaked state: a 20 x 20 px element somewhere else, far brighter
    let mut moved = a.clone();
    for px in 0..400 {
        moved[3 * px] = 250;
    }
    assert!(pixels_differing(&a, &moved, NOISE_LEVEL) > MAX_CONTENT_PIXELS);
}
