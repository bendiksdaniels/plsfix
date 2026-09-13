// Reconciles the saved workbook against the builders' tally and the layout:
// sheet order and the hidden sheet, per-sheet cell and formula counts, the
// planted errors, the defined names and the chart ranges.

use std::io::{Cursor, Read};

use smt_demo::layout::{cell, cell_abs, last_year_col, pnl, variance as vlayout, year_col, FIRST_YEAR_COL, PNL_SHEET};
use smt_demo::sheets::{
    pnl::{planted_cells, table_range},
    rounding, scratch, variance, BROKEN_NAME, SHEETS,
};
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

/// The workbook's shared string table, in index order: `t="s"` cells
/// reference it by position, so guide text (always a shared string) needs
/// this to read back as more than an opaque index.
fn shared_strings(archive: &mut Archive) -> Vec<String> {
    let xml = entry(archive, "xl/sharedStrings.xml");
    let mut strings = Vec::new();
    let mut rest = xml.as_str();
    while let Some(start) = rest.find("<t") {
        let tag_end = rest[start..].find('>').expect("tag close") + start;
        let text_start = tag_end + 1;
        let text_end = rest[text_start..].find("</t>").expect("text end") + text_start;
        strings.push(unescape_xml(&rest[text_start..text_end]));
        rest = &rest[text_end + "</t>".len()..];
    }
    strings
}

/// Undoes the three entities `rust_xlsxwriter` escapes in text data (it does
/// not escape quotes, unlike attribute values).
fn unescape_xml(text: &str) -> String {
    text.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
}

/// The text of a `t="s"` (shared string) cell, resolved through `strings`.
fn cell_text(sheet: &str, address: &str, strings: &[String]) -> String {
    let xml = cell_xml(sheet, address);
    let start = xml.find("<v>").unwrap_or_else(|| panic!("{address}: not a string cell: {xml}")) + "<v>".len();
    let end = xml[start..].find("</v>").expect("value end") + start;
    let index: usize = xml[start..end].parse().expect("shared string index");
    strings[index].clone()
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
    // Chart ranges are computed from the layout constants, not copied by hand,
    // so a future row shift cannot leave a stale address behind here again.
    let revenue_range = format!("{}:{}", cell_abs(pnl::REVENUE, FIRST_YEAR_COL), cell_abs(pnl::REVENUE, last_year_col()));
    let revenue = entry(&mut archive, "xl/charts/chart1.xml");
    assert!(revenue.contains(&format!("'P&amp;L'!{revenue_range}")), "revenue values: {revenue}");
    let margin_range = format!(
        "{}:{}",
        cell_abs(pnl::EBITDA_MARGIN, FIRST_YEAR_COL),
        cell_abs(pnl::EBITDA_MARGIN, last_year_col())
    );
    let margin = entry(&mut archive, "xl/charts/chart2.xml");
    assert!(margin.contains(&format!("'P&amp;L'!{margin_range}")), "margin values: {margin}");
    let pie = entry(&mut archive, "xl/charts/chart3.xml");
    let pie_range = format!("{}!{}", rounding::NAME, rounding::pie_values_address_abs());
    assert!(pie.contains("<c:pieChart>") && pie.contains(&pie_range), "pie: {pie}");
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

/// True when `address` has no `<c r="address" ...>` element at all: an
/// empty cell, never written.
fn cell_is_absent(sheet: &str, address: &str) -> bool {
    !sheet.contains(&format!("<c r=\"{address}\""))
}

// Every visible sheet opens with the 7-row guide band (GUIDE_ROWS in
// layout.rs): title at A1, the Commands line always at A6 (bands never
// exceed four tasks, so A5 is spare here), a blank spacer at A7, then the
// sheet's own first row of real content at A8. Scratch is hidden on
// purpose and carries no band.
#[test]
fn every_visible_sheet_opens_with_its_guide() {
    let (mut archive, _) = archive();
    let strings = shared_strings(&mut archive);
    for (name, _) in SHEETS {
        if name == scratch::NAME {
            continue;
        }
        let sheet = sheet_xml(&mut archive, name);
        let title = cell_text(&sheet, "A1", &strings);
        assert!(title.starts_with("Try on this sheet"), "{name}: A1 = {title:?}");
        let commands = cell_text(&sheet, "A6", &strings);
        assert!(commands.starts_with("Commands:"), "{name}: A6 = {commands:?}");
        assert!(cell_is_absent(&sheet, "A5"), "{name}: A5 should be spare (no sheet uses four tasks yet)");
        assert!(cell_is_absent(&sheet, "A7"), "{name}: A7 should be the band's blank spacer row");
        let own_title = cell_text(&sheet, "A8", &strings);
        assert!(
            !own_title.is_empty() && !own_title.starts_with("Try on this sheet"),
            "{name}: A8 should be the sheet's own title, was {own_title:?}"
        );
    }
}

/// Each `(name, key)` a Commands line cites, split on the unambiguous
/// `"), "` between entries (no real action name or key contains a
/// parenthesis or that exact sequence) and then on the first `" ("` inside
/// each entry, so a comma inside a name (e.g. "pls,fix") is never mistaken
/// for the entry separator.
fn cited_commands(commands_line: &str) -> Vec<(String, String)> {
    let body = commands_line.strip_prefix("Commands: ").unwrap_or(commands_line);
    body.split("), ")
        .map(|entry| {
            let entry = entry.trim_end_matches(')');
            let (name, key) = entry.split_once(" (").unwrap_or_else(|| panic!("malformed command entry {entry:?}"));
            (name.to_string(), key.to_string())
        })
        .collect()
}

/// Every quoted value following `needle` (e.g. `"\"id\": \""`), in file
/// order: one flat field of every object in an array, read without a JSON
/// dependency.
fn extract_quoted(json: &str, needle: &str) -> Vec<String> {
    json.match_indices(needle)
        .map(|(start, _)| {
            let after = &json[start + needle.len()..];
            let end = after.find('"').expect("closing quote");
            after[..end].to_string()
        })
        .collect()
}

/// `(name, key)` for every real action: `public/shortcuts.json`'s
/// `actions[].id -> name` joined with `shortcuts[].action -> key.default`.
/// The only names and keys a Commands line may ever cite; extracted by
/// hand instead of a JSON dependency, since both arrays are simple flat
/// objects in a fixed field order.
fn real_action_keys() -> Vec<(String, String)> {
    let json = std::fs::read_to_string("../public/shortcuts.json").expect("shortcuts.json");
    let ids = extract_quoted(&json, "\"id\": \"");
    let names = extract_quoted(&json, "\"name\": \"");
    let shortcut_ids = extract_quoted(&json, "\"action\": \"");
    let keys = extract_quoted(&json, "\"default\": \"");
    assert_eq!(ids.len(), names.len(), "actions[]: id/name count mismatch");
    assert_eq!(shortcut_ids.len(), keys.len(), "shortcuts[]: action/key count mismatch");
    ids.into_iter()
        .zip(names)
        .map(|(id, name)| {
            let position = shortcut_ids.iter().position(|action_id| *action_id == id);
            let key = position.unwrap_or_else(|| panic!("{id}: no shortcuts[] entry"));
            (name, keys[key].clone())
        })
        .collect()
}

#[test]
fn guide_names_only_real_commands() {
    let (mut archive, _) = archive();
    let strings = shared_strings(&mut archive);
    let real = real_action_keys();
    for (name, _) in SHEETS {
        if name == scratch::NAME {
            continue;
        }
        let sheet = sheet_xml(&mut archive, name);
        let commands = cell_text(&sheet, "A6", &strings);
        for (cited_name, cited_key) in cited_commands(&commands) {
            let real_key = real
                .iter()
                .find(|(real_name, _)| *real_name == cited_name)
                .unwrap_or_else(|| panic!("{name}: unknown command {cited_name:?} in {commands:?}"));
            assert_eq!(
                cited_key, real_key.1,
                "{name}: {cited_name} cites key {cited_key:?}, shortcuts.json says {:?}",
                real_key.1
            );
        }
    }
}

// The P&L guide's third task must name the real, currently-computed export
// range, not a literal that could go stale on the next row shift.
#[test]
fn pnl_guide_names_the_real_export_range() {
    let (mut archive, _) = archive();
    let strings = shared_strings(&mut archive);
    let sheet = sheet_xml(&mut archive, PNL_SHEET);
    let task3 = cell_text(&sheet, "A4", &strings);
    let expected = table_range();
    assert!(task3.contains(&expected), "P&L task 3 = {task3:?}, expected the range {expected}");
}
