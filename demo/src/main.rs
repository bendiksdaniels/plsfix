// Entry point: builds the demo workbook and saves it, by default to
// demo/out/Demo Model.xlsx, or to the path given as the first argument.
// Prints one line per sheet with what was written, then the total.

use std::path::{Path, PathBuf};

use smt_demo::tally::Tally;

const OUT_DIR: &str = "out";
const OUT_FILE: &str = "Demo Model.xlsx";

fn default_path() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(OUT_DIR).join(OUT_FILE)
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = std::env::args().nth(1).map_or_else(default_path, PathBuf::from);
    let (mut workbook, tallies) = smt_demo::build()?;
    if let Some(dir) = path.parent() {
        std::fs::create_dir_all(dir)?;
    }
    workbook.save(&path)?;
    let mut total = Tally::default();
    for (name, tally) in &tallies {
        println!(
            "{name}: {} labels, {} numbers, {} formulas",
            tally.strings, tally.numbers, tally.formulas
        );
        total += *tally;
    }
    println!(
        "saved {} sheets, {} cells -> {}",
        tallies.len(),
        total.cells(),
        path.display()
    );
    Ok(())
}
