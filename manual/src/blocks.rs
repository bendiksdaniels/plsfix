// One content block turned into Word elements: paragraphs, lists, tables and
// pictures. Owns the numbering instances and the alt-text order; `render.rs`
// hands it a `Ctx` and gets the document back with the block appended.

use docx_rs::*;

use crate::content::Block;
use crate::meta;
use crate::shots;
use crate::style;

/// The abstract numbering ids the whole document shares.
const DECIMAL_LIST: usize = 1;
const BULLET_LIST: usize = 2;
/// Numbering instance ids start above the two abstract ones.
const FIRST_NUM_ID: usize = 10;
/// A4 minus the two margins, in twips.
const CONTENT_WIDTH: usize = 9638;

pub struct Ctx {
    pub version: String,
    /// Alt text for each picture, in the order the pictures are added.
    pub alts: Vec<String>,
    /// Numbering instances to register on the document once it is built.
    pub numberings: Vec<Numbering>,
}

impl Ctx {
    pub fn new(version: String) -> Self {
        Ctx {
            version,
            alts: Vec::new(),
            numberings: Vec::new(),
        }
    }

    /// A fresh list instance, so every list restarts at one.
    fn list(&mut self, abstract_id: usize) -> usize {
        let id = FIRST_NUM_ID + self.numberings.len();
        self.numberings
            .push(Numbering::new(id, abstract_id).add_override(LevelOverride::new(0).start(1)));
        id
    }

    /// Substitutes the placeholders content may carry.
    fn fill(&self, value: &str) -> String {
        value.replace("{version}", &self.version)
    }
}

pub fn block(docx: Docx, ctx: &mut Ctx, item: &Block) -> Docx {
    match item {
        Block::Para(text) => docx.add_paragraph(paragraph(&ctx.fill(text))),
        Block::Steps(items) => list(docx, ctx, items, DECIMAL_LIST),
        Block::Bullets(items) => list(docx, ctx, items, BULLET_LIST),
        Block::Table { head, rows } => {
            let body: Vec<Vec<String>> = rows
                .iter()
                .map(|row| row.iter().map(|c| ctx.fill(c)).collect())
                .collect();
            table(docx, head, &body)
        }
        Block::Shortcuts => shortcut_table(docx),
        Block::Image { file, alt } => image(docx, ctx, file, alt),
    }
}

pub fn paragraph(text: &str) -> Paragraph {
    Paragraph::new()
        .add_run(style::text(text))
        .line_spacing(LineSpacing::new().after(120).line(276))
}

fn list(docx: Docx, ctx: &mut Ctx, items: &[&str], abstract_id: usize) -> Docx {
    let id = ctx.list(abstract_id);
    items.iter().fold(docx, |doc, item| {
        doc.add_paragraph(
            Paragraph::new()
                .add_run(style::text(&ctx.fill(item)))
                .numbering(NumberingId::new(id), IndentLevel::new(0))
                .line_spacing(LineSpacing::new().after(60).line(276)),
        )
    })
}

fn table(docx: Docx, head: &[&str], rows: &[Vec<String>]) -> Docx {
    let grid = grid(head.len());
    let mut table_rows = vec![
        TableRow::new(
            head.iter()
                .enumerate()
                .map(|(i, text)| head_cell(text, grid[i]))
                .collect(),
        )
        .cant_split(),
    ];
    for row in rows {
        table_rows.push(TableRow::new(
            row.iter()
                .enumerate()
                .map(|(i, text)| body_cell(text, grid[i]))
                .collect(),
        ));
    }
    docx.add_table(
        Table::new(table_rows)
            .set_grid(grid)
            .layout(TableLayoutType::Fixed)
            .width(CONTENT_WIDTH, WidthType::Dxa)
            .set_borders(style::table_rules())
            .margins(style::table_margins()),
    )
    .add_paragraph(Paragraph::new())
}

/// Column widths in twips: the first column carries the label, the rest split
/// what is left evenly.
fn grid(columns: usize) -> Vec<usize> {
    if columns < 2 {
        return vec![CONTENT_WIDTH];
    }
    let first = CONTENT_WIDTH * 28 / 100;
    let rest = (CONTENT_WIDTH - first) / (columns - 1);
    let mut widths = vec![first];
    widths.extend(std::iter::repeat_n(rest, columns - 1));
    widths
}

fn head_cell(text: &str, width: usize) -> TableCell {
    TableCell::new()
        .width(width, WidthType::Dxa)
        .shading(Shading::new().fill(style::MINT_TINT))
        .add_paragraph(
            Paragraph::new().add_run(
                style::tinted(text, style::TABLE, style::NAVY)
                    .bold()
                    .fonts(style::fonts()),
            ),
        )
}

fn body_cell(text: &str, width: usize) -> TableCell {
    TableCell::new()
        .width(width, WidthType::Dxa)
        .add_paragraph(Paragraph::new().add_run(style::tinted(text, style::TABLE, style::INK)))
}

fn shortcut_table(docx: Docx) -> Docx {
    let rows: Vec<Vec<String>> = meta::shortcuts()
        .iter()
        .map(|s| {
            vec![
                format!("\"{}\"", s.name),
                s.keys.clone(),
                crate::content::reference::gloss(&s.id).to_string(),
            ]
        })
        .collect();
    table(docx, &["Darbība", "Taustiņi", "Ko tā dara"], &rows)
}

fn image(docx: Docx, ctx: &mut Ctx, file: &str, alt: &str) -> Docx {
    let shot = shots::load(file);
    ctx.alts.push(alt.to_string());
    docx.add_paragraph(picture(&shot, shot.size))
        .add_paragraph(caption(alt))
}

pub fn picture(shot: &shots::Shot, size: (u32, u32)) -> Paragraph {
    let pic = Pic::new(&shot.bytes).size(size.0, size.1);
    Paragraph::new()
        .add_run(Run::new().add_image(pic))
        .align(AlignmentType::Center)
}

fn caption(text: &str) -> Paragraph {
    Paragraph::new()
        .add_run(style::tinted(text, style::SMALL, style::GREY).italic())
        .align(AlignmentType::Center)
        .line_spacing(LineSpacing::new().after(200))
}

/// The two list shapes every list instance is based on.
pub fn abstract_numberings() -> Vec<AbstractNumbering> {
    vec![
        AbstractNumbering::new(DECIMAL_LIST).add_level(level("decimal", "%1.")),
        AbstractNumbering::new(BULLET_LIST).add_level(level("bullet", "\u{2022}")),
    ]
}

fn level(format: &str, text: &str) -> Level {
    Level::new(
        0,
        Start::new(1),
        NumberFormat::new(format),
        LevelText::new(text),
        LevelJc::new("left"),
    )
    .indent(Some(420), Some(SpecialIndentType::Hanging(360)), None, None)
}
