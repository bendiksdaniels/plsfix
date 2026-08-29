// The document's look in one place: page geometry, brand colours, type sizes
// and the run helpers every renderer uses. Owns every magic number in the
// layout; nothing else in the crate may hardcode a colour or a size.

use docx_rs::*;

/// Brand navy, used for every heading.
pub const NAVY: &str = "14213D";
/// Brand mint, used for the accents beside the navy.
pub const MINT: &str = "2EC4B6";
/// A mint tint light enough to read black text on, for table header rows.
pub const MINT_TINT: &str = "DFF5F2";
/// Body text.
pub const INK: &str = "1A1A1A";
/// Secondary text: captions and the footer.
pub const GREY: &str = "5B6472";
/// Table rules, light enough not to fight the text.
pub const RULE: &str = "C8CEDA";

/// A4 in twips: 210 x 297 mm.
pub const PAGE_W: u32 = 11906;
pub const PAGE_H: u32 = 16838;
/// 2 cm margins, and the footer 1.25 cm from the paper edge.
pub const MARGIN: i32 = 1134;
pub const FOOTER_MARGIN: i32 = 709;

/// Half-points, the unit Word uses for font size.
pub const BODY: usize = 22;
pub const SMALL: usize = 18;
pub const TABLE: usize = 20;
pub const H1: usize = 36;
pub const H2: usize = 26;
pub const COVER_TITLE: usize = 72;
pub const COVER_SUB: usize = 28;

/// EMU per centimetre, the unit Word uses for picture size.
pub const EMU_PER_CM: f64 = 360_000.0;
/// No picture is wider than this, and none is taller.
pub const MAX_IMAGE_W_CM: f64 = 15.0;
pub const MAX_IMAGE_H_CM: f64 = 11.0;

/// The body font. Word substitutes Arial where Aptos is missing, which
/// `postprocess::add_font_fallback` writes into the font table.
pub const FONT: &str = "Aptos";

pub fn fonts() -> RunFonts {
    RunFonts::new()
        .ascii(FONT)
        .hi_ansi(FONT)
        .east_asia(FONT)
        .cs(FONT)
}

/// Body text in the default colour and size.
pub fn text(value: &str) -> Run {
    Run::new()
        .fonts(fonts())
        .size(BODY)
        .color(INK)
        .add_text(value)
}

/// A run at a given size and colour, for headings and captions.
pub fn tinted(value: &str, size: usize, color: &str) -> Run {
    Run::new()
        .fonts(fonts())
        .size(size)
        .color(color)
        .add_text(value)
}

/// Thin grey rules on every side of a table and between its cells.
pub fn table_rules() -> TableBorders {
    [
        TableBorderPosition::Top,
        TableBorderPosition::Left,
        TableBorderPosition::Bottom,
        TableBorderPosition::Right,
        TableBorderPosition::InsideH,
        TableBorderPosition::InsideV,
    ]
    .into_iter()
    .fold(TableBorders::with_empty(), |borders, position| {
        borders.set(
            TableBorder::new(position)
                .border_type(BorderType::Single)
                .size(2)
                .color(RULE),
        )
    })
}

/// Breathing room inside every table cell, in twips.
pub fn table_margins() -> TableCellMargins {
    TableCellMargins::new().margin(60, 108, 60, 108)
}

/// The mint rule that sits under every chapter heading.
pub fn heading_rule() -> ParagraphBorders {
    ParagraphBorders::with_empty().set(
        ParagraphBorder::new(ParagraphBorderPosition::Bottom)
            .val(BorderType::Single)
            .size(8)
            .space(6)
            .color(MINT),
    )
}
