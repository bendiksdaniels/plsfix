// Brand colours and the handful of cell formats the demo sheets share. Fonts
// keep Excel's default colour on purpose: Autocolor is meant to have work to do.

use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder};

pub const NAVY: Color = Color::RGB(0x14_21_3D);
pub const TEAL: Color = Color::RGB(0x2E_C4_B6);
pub const MINT_TINT: Color = Color::RGB(0xE6_F8_F6);
pub const EUR_K: &str = "#,##0;(#,##0);\"-\"";
pub const PERCENT: &str = "0.0%";
pub const PERCENT_2: &str = "0.00%";
pub const POINTS: &str = "0.00";
pub const WHOLE: &str = "0";
pub const NUMBER_1DP: &str = "#,##0.0";
const TITLE_SIZE: f64 = 14.0;

pub struct Styles {
    pub title: Format,
    pub note: Format,
    pub header: Format,
    pub header_left: Format,
    pub label: Format,
    pub label_bold: Format,
    pub eur: Format,
    pub eur_bold: Format,
    pub pct: Format,
    pub pct2: Format,
    pub points: Format,
    pub whole: Format,
    pub number_1dp: Format,
    pub plain: Format,
    pub guide_title: Format,
    pub guide_text: Format,
}

impl Styles {
    pub fn new() -> Styles {
        Styles {
            title: Format::new().set_bold().set_font_size(TITLE_SIZE).set_font_color(NAVY),
            note: Format::new().set_italic().set_font_color(TEAL),
            header: Format::new()
                .set_bold()
                .set_align(FormatAlign::Right)
                .set_border_bottom(FormatBorder::Thin),
            header_left: Format::new().set_bold().set_border_bottom(FormatBorder::Thin),
            label: Format::new(),
            label_bold: Format::new().set_bold(),
            eur: Format::new().set_num_format(EUR_K),
            eur_bold: Format::new().set_bold().set_num_format(EUR_K),
            pct: Format::new().set_num_format(PERCENT),
            pct2: Format::new().set_num_format(PERCENT_2),
            points: Format::new().set_num_format(POINTS),
            whole: Format::new().set_num_format(WHOLE),
            number_1dp: Format::new().set_num_format(NUMBER_1DP),
            plain: Format::new(),
            // No wrap: column A is as narrow as 4 characters on some sheets
            // (Start here), so wrapped text broke into unreadable slivers.
            // The line overflows across the empty cells to its right instead
            // (rows 1-7 hold column A only, everywhere); set_row_format
            // paints the tint over those cells since they carry no format
            // of their own.
            guide_title: Format::new().set_bold().set_font_color(NAVY).set_background_color(MINT_TINT),
            guide_text: Format::new().set_background_color(MINT_TINT),
        }
    }
}

impl Default for Styles {
    fn default() -> Self {
        Self::new()
    }
}
