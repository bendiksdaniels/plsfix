// Reconciles the saved workbook against the builders' tally and the layout:
// sheet order and the hidden sheet, per-sheet cell and formula counts, the
// planted errors, the defined names and the chart ranges.

use std::io::{Cursor, Read};

use smt_demo::layout::{cell, pnl, variance as vlayout, year_col, PNL_SHEET};
use smt_demo::sheets::{pnl::planted_cells, scratch, variance, BROKEN_NAME, SHEETS};
use smt_demo::tally::Tally;

type Archive = zip::ZipArchive<Cursor<Vec<u8>>>;

fn archive() -> (Archive, Vec<(&'static str, Tally)>) {
    let (mut workbook, tallies) = smt_demo::build().expect("build");
    let bytes = workbook.save_to_buffer().expect("save");
    (zip::ZipArchive::new(Cursor::new(bytes)).expect("zip"), tallies)
}

fn entry(archive: &mut Archive, name: &str) -> String {
    let mut text = String::new();
    archive.by_name(name).expect(name).read_to_string(&mut text).expect("utf8");
    text
}

fn sheet_xml(archive: &mut Archive, name: &str) -> String {
    let index = SHEETS.iter().position(|(sheet, _)| *sheet == name).expect(name);
    entry(archive, &format!("xl/worksheets/sheet{}.xml", index + 1))
}

/// The `<c ...>...</c>` element of one cell.
fn cell_xml(sheet: &str, address: &str) -> String {
    let start = sheet.find(&format!("<c r=\"{address}\"")).unwrap_or_else(|| panic!("cell {address}"));
    let rest = &sheet[start..];
    let end = rest.find("</c>").expect("cell end");
    rest[..end].to_string()
}

#[test]
fn sheets_are_saved_in_order_and_scratch_is_hidden() {
    let (mut archive, _) = archive();
    let workbook = entry(&mut archive, "xl/workbook.xml");
    let mut last = 0;
    for (name, _) in SHEETS {
        let needle = format!("name=\"{}\"", name.replace('&', "&amp;"));
        let at = workbook.find(&needle).unwrap_or_else(|| panic!("sheet {name} missing"));
        assert!(at > last, "{name} out of order");
        last = at;
    }
    let scratch = workbook.find(&format!("name=\"{}\"", scratch::NAME)).expect("scratch");
    let tail = &workbook[..scratch + 200];
    assert!(tail.contains("state=\"hidden\""), "Scratch must be hidden");
}

#[test]
fn every_written_cell_is_in_the_file() {
    let (mut archive, tallies) = archive();
    assert_eq!(tallies.len(), SHEETS.len());
    for (name, tally) in &tallies {
        let sheet = sheet_xml(&mut archive, name);
        assert_eq!(sheet.matches("<c ").count(), tally.cells(), "{name}: cells");
        assert_eq!(sheet.matches("<f>").count(), tally.formulas, "{name}: formulas");
        assert_eq!(sheet.matches(" t=\"s\"").count(), tally.strings, "{name}: strings");
        assert!(tally.cells() > 0, "{name}: empty");
    }
}

#[test]
fn planted_errors_are_in_the_pnl() {
    let (mut archive, _) = archive();
    let sheet = sheet_xml(&mut archive, PNL_SHEET);
    let planted = planted_cells();
    let hardcode = cell_xml(&sheet, &planted[0]);
    let expected = format!("<f>-{}*0.21</f>", cell(pnl::REVENUE, year_col(3)));
    assert!(hardcode.contains(&expected), "{hardcode}");
    let typed = cell_xml(&sheet, &planted[1]);
    assert!(!typed.contains("<f>"), "typed number must not be a formula: {typed}");
    assert!(typed.contains("<v>-560</v>"), "{typed}");
}

#[test]
fn names_and_charts_are_defined() {
    let (mut archive, _) = archive();
    let workbook = entry(&mut archive, "xl/workbook.xml");
    for name in ["Revenue_2024A", "TaxRate_2025E", BROKEN_NAME, scratch::AREA_NAME] {
        assert!(workbook.contains(&format!("name=\"{name}\"")), "{name}");
    }
    assert!(workbook.contains(&scratch::area_formula()[1..]), "Scratch_area formula");
    assert!(workbook.contains(&format!("name=\"{BROKEN_NAME}\">#REF!<")), "broken name");
    let revenue = entry(&mut archive, "xl/charts/chart1.xml");
    assert!(revenue.contains("'P&amp;L'!$C$4:$H$4"), "revenue values: {revenue}");
    let margin = entry(&mut archive, "xl/charts/chart2.xml");
    assert!(margin.contains("'P&amp;L'!$C$11:$H$11"), "margin values: {margin}");
    let pie = entry(&mut archive, "xl/charts/chart3.xml");
    assert!(pie.contains("<c:pieChart>") && pie.contains("Rounding!$B$4:$B$8"), "pie: {pie}");
}

// Every chart in the demo labels its points with the value and nothing else,
// the same rule the add-in applies to the charts it builds and restyles.
#[test]
fn every_chart_labels_the_values_only() {
    let (mut archive, _) = archive();
    for index in 1..=3 {
        let chart = entry(&mut archive, &format!("xl/charts/chart{index}.xml"));
        assert!(chart.contains("<c:showVal val=\"1\"/>"), "chart {index} shows values: {chart}");
        for part in ["showPercent", "showCatName", "showSerName", "showLegendKey"] {
            let on = format!("<c:{part} val=\"1\"/>");
            assert!(!chart.contains(&on), "chart {index} shows {part}: {chart}");
        }
    }
}

#[test]
fn variance_amounts_match_the_saved_cells() {
    let (mut archive, _) = archive();
    let sheet = sheet_xml(&mut archive, variance::NAME);
    for (i, (label, amount)) in variance::AMOUNTS.iter().enumerate() {
        let address = cell(vlayout::FIRST_ROW + i as u32, vlayout::AMOUNT_COL);
        let xml = cell_xml(&sheet, &address);
        assert!(!xml.contains("<f>"), "{label} must be a literal number: {xml}");
        assert!(xml.contains(&format!("<v>{}</v>", *amount as i64)), "{label}: {xml}");
    }
}

#[test]
fn variance_target_lands_at_its_advertised_cell() {
    let (mut archive, _) = archive();
    let sheet = sheet_xml(&mut archive, variance::NAME);
    let address = cell(vlayout::TARGET_ROW, vlayout::AMOUNT_COL);
    let xml = cell_xml(&sheet, &address);
    assert!(!xml.contains("<f>"), "target must be a literal number: {xml}");
    assert!(xml.contains(&format!("<v>{}</v>", variance::TARGET as i64)), "{xml}");
}

// The puzzle only works as "Find a combination" if exactly one combination of
// invoice lines reaches the target: proved here by brute force (2^14 - 1
// non-empty subsets) instead of trusted by eye.
#[test]
fn variance_target_is_reachable_by_exactly_one_subset() {
    let amounts: Vec<f64> = variance::AMOUNTS.iter().map(|(_, value)| *value).collect();
    let (hits, size) = subset_sum_hits(&amounts, variance::TARGET);
    assert_eq!(hits, 1, "expected exactly one subset of the amounts to reach the target");
    assert!((3..=4).contains(&size), "the one subset should be 3-4 cells, was {size}");
}

/// How many non-empty subsets of `values` sum to `target`, and the cell
/// count of the last one found (meaningful only when there is one hit).
fn subset_sum_hits(values: &[f64], target: f64) -> (u32, u32) {
    let mut hits = 0;
    let mut size = 0;
    for mask in 1u32..(1 << values.len()) {
        let sum: f64 = (0..values.len())
            .filter(|bit| mask & (1 << bit) != 0)
            .map(|bit| values[bit])
            .sum();
        if (sum - target).abs() < 1e-6 {
            hits += 1;
            size = mask.count_ones();
        }
    }
    (hits, size)
}
