// Entry point: builds the manual and writes it to manual/out. Owns nothing but
// the file write and the one line of output that names what was produced.

fn main() -> std::io::Result<()> {
    let bytes = manual::render::build();
    let path = manual::output_path();
    if let Some(folder) = path.parent() {
        std::fs::create_dir_all(folder)?;
    }
    std::fs::write(&path, &bytes)?;
    println!(
        "manual {}: {} ({} KB)",
        manual::meta::version(),
        path.display(),
        bytes.len() / 1024
    );
    Ok(())
}
