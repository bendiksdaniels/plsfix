// The pls,fix user manual generator: content as data, rendered into a Word
// document. Owns the crate's public surface, `render::build` and the output
// path; everything else is a module behind it.

pub mod blocks;
pub mod content;
pub mod meta;
pub mod postprocess;
pub mod render;
pub mod shots;
pub mod style;

/// Where the finished manual is written.
pub fn output_path() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("out")
        .join("pls,fix rokasgrāmata.docx")
}
