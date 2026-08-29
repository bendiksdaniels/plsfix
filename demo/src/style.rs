// Brand colours and the handful of cell formats the demo sheets share. Fonts
// keep Excel's default colour on purpose: Autocolor is meant to have work to do.

use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder};

pub const CHARCOAL: Color = Color::RGB(0x28_26_23);
pub const BRONZE: Color = Color::RGB(0xB2_7E_54);
pub const EUR_K: &str = "#,##0;(#,##0);\"-\"";
pub const PERCENT: &str = "0.0%";
pub const PERCENT_2: &str = "0.00%";
pub const POINTS: &str = "0.00";
pub const WHOLE: &str = "0";
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
    pub plain: Format,
}

impl Styles {
    pub fn new() -> Styles {
        Styles {
            title: Format::new().set_bold().set_font_size(TITLE_SIZE).set_font_color(CHARCOAL),
            note: Format::new().set_italic().set_font_color(BRONZE),
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
            plain: Format::new(),
        }
    }
}

impl Default for Styles {
    fn default() -> Self {
        Self::new()
    }
}
