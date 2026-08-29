// Turns the content model into a Word document: page setup, cover, table of
// contents, headings and the footer that repeats on every page. Owns the
// document's shape; the words come from `content`, the look from `style`.

use docx_rs::*;

use crate::blocks::{self, Ctx};
use crate::content;
use crate::meta;
use crate::postprocess;
use crate::shots;
use crate::style;

/// The finished .docx, ready to write to disk.
pub fn build() -> Vec<u8> {
    let mut ctx = Ctx::new(meta::version());
    let mut packed = Vec::new();
    document(&mut ctx)
        .build()
        .pack(std::io::Cursor::new(&mut packed))
        .expect("pack the docx");
    postprocess::finish(&packed, &ctx.alts)
}

fn document(ctx: &mut Ctx) -> Docx {
    let mut docx = chapters(contents(cover(base(ctx), ctx)), ctx);
    for numbering in std::mem::take(&mut ctx.numberings) {
        docx = docx.add_numbering(numbering);
    }
    docx
}

fn base(ctx: &Ctx) -> Docx {
    let mut docx = Docx::new()
        .page_size(style::PAGE_W, style::PAGE_H)
        .page_margin(page_margin())
        .default_fonts(style::fonts())
        .default_size(style::BODY)
        .footer(footer(&ctx.version))
        .add_style(heading_style("Heading1", "Heading 1", style::H1, 0))
        .add_style(heading_style("Heading2", "Heading 2", style::H2, 1));
    for numbering in blocks::abstract_numberings() {
        docx = docx.add_abstract_numbering(numbering);
    }
    docx
}

fn page_margin() -> PageMargin {
    PageMargin::new()
        .top(style::MARGIN)
        .bottom(style::MARGIN)
        .left(style::MARGIN)
        .right(style::MARGIN)
        .header(style::FOOTER_MARGIN)
        .footer(style::FOOTER_MARGIN)
}

fn heading_style(id: &str, name: &str, size: usize, level: usize) -> Style {
    Style::new(id, StyleType::Paragraph)
        .name(name)
        .size(size)
        .color(style::NAVY)
        .bold()
        .fonts(style::fonts())
        .outline_lvl(level)
}

/// The same footer on every page, the title page included: version on the left
/// of the page number, both centred.
fn footer(version: &str) -> Footer {
    Footer::new().add_paragraph(
        Paragraph::new()
            .align(AlignmentType::Center)
            .add_run(style::tinted(
                &format!("pls,fix rokasgrāmata {version}   |   "),
                style::SMALL,
                style::GREY,
            ))
            .add_run(page_number()),
    )
}

fn page_number() -> Run {
    Run::new()
        .fonts(style::fonts())
        .size(style::SMALL)
        .color(style::GREY)
        .add_field_char(FieldCharType::Begin, false)
        .add_instr_text(InstrText::PAGE(InstrPAGE {}))
        .add_field_char(FieldCharType::Separate, false)
        .add_text("1")
        .add_field_char(FieldCharType::End, false)
}

fn cover(docx: Docx, ctx: &mut Ctx) -> Docx {
    let shot = shots::load(content::COVER_IMAGE);
    ctx.alts.push(content::COVER_ALT.to_string());
    docx.add_paragraph(Paragraph::new())
        .add_paragraph(blocks::picture(&shot, shots::cover_size(&shot)))
        .add_paragraph(cover_line("pls,fix", style::COVER_TITLE, style::NAVY))
        .add_paragraph(cover_line(
            "Lietotāja rokasgrāmata",
            style::COVER_SUB,
            style::MINT,
        ))
        .add_paragraph(cover_line(
            "Excel un PowerPoint pievienojumprogramma finanšu modelēšanai",
            style::BODY,
            style::INK,
        ))
        .add_paragraph(cover_line(&ctx.version.clone(), style::BODY, style::GREY))
}

fn cover_line(text: &str, size: usize, color: &str) -> Paragraph {
    Paragraph::new()
        .add_run(style::tinted(text, size, color).bold())
        .align(AlignmentType::Center)
        .line_spacing(LineSpacing::new().after(160))
}

/// The contents page. Its own title is a plain paragraph, not a heading, so it
/// does not list itself.
fn contents(docx: Docx) -> Docx {
    docx.add_paragraph(
        Paragraph::new()
            .page_break_before(true)
            .keep_next(true)
            .line_spacing(LineSpacing::new().after(200))
            .add_run(style::tinted("Saturs", style::H1, style::NAVY).bold()),
    )
    .add_table_of_contents(
        TableOfContents::new()
            .heading_styles_range(1, 2)
            .alias("Saturs")
            .auto()
            .dirty(),
    )
}

fn chapters(docx: Docx, ctx: &mut Ctx) -> Docx {
    let mut out = docx;
    for (index, chapter) in content::chapters().into_iter().enumerate() {
        out = out.add_paragraph(heading1(index + 1, chapter.title));
        for section in chapter.sections {
            out = out.add_paragraph(heading2(section.title));
            for item in &section.blocks {
                out = blocks::block(out, ctx, item);
            }
        }
    }
    out
}

fn heading1(number: usize, title: &str) -> Paragraph {
    let mut paragraph = Paragraph::new()
        .style("Heading1")
        .page_break_before(true)
        .keep_next(true)
        .add_run(style::tinted(&format!("{number}  "), style::H1, style::MINT).bold())
        .add_run(style::tinted(title, style::H1, style::NAVY).bold())
        .line_spacing(LineSpacing::new().before(120).after(200));
    paragraph.property = paragraph.property.set_borders(style::heading_rule());
    paragraph
}

fn heading2(title: &str) -> Paragraph {
    Paragraph::new()
        .style("Heading2")
        .keep_next(true)
        .add_run(style::tinted(title, style::H2, style::NAVY).bold())
        .line_spacing(LineSpacing::new().before(240).after(100))
}
