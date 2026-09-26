//! wav.rs tests: the header says 32-bit float stereo at the given rate and the sizes add up.
use plsfix_video::core::sound::bus::Bus;
use plsfix_video::io::wav::write;

#[test]
fn header_is_float_stereo_and_sizes_add_up() {
    let mut bus = Bus::new(10);
    bus.add(3, 0.5, -0.5);
    let p = std::env::temp_dir().join(format!("pv-wav-{}.wav", std::process::id()));
    write(&p, &bus, 48_000).unwrap();
    let b = std::fs::read(&p).unwrap();
    assert_eq!(&b[0..4], b"RIFF");
    assert_eq!(u32::from_le_bytes(b[4..8].try_into().unwrap()) as usize, b.len() - 8);
    assert_eq!(u16::from_le_bytes([b[20], b[21]]), 3, "IEEE float");
    assert_eq!(u16::from_le_bytes([b[22], b[23]]), 2, "stereo");
    assert_eq!(u32::from_le_bytes(b[24..28].try_into().unwrap()), 48_000);
    assert_eq!(b.len(), 58 + 10 * 8);
    std::fs::remove_file(&p).ok();
}
