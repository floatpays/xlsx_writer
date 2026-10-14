//! Building rust_xlsxwriter `Format`s from the decoded format options.

use rust_xlsxwriter::{
    Color, Format, FormatAlign, FormatBorder, FormatPattern, FormatScript, FormatUnderline,
};

use crate::error::{WriteError, WriteResult};
use crate::protocol::{
    BorderStyle, CellAlignPos, CellData, CellFormat, CellPattern, CellVAlignPos, UnderlineStyle,
};

/// Excel's maximum font size, in points.
const MAX_FONT_SIZE: f64 = 409.0;
/// Excel's maximum indent level.
const MAX_INDENT: u8 = 250;

const DEFAULT_DATE_FORMAT: &str = "yyyy-mm-dd";
const DEFAULT_DATE_TIME_FORMAT: &str = "yyyy-mm-ddThh:mm:ss";

/// The format a cell is written with: `None` when the cell has no format of
/// its own, so that rust_xlsxwriter applies its defaults (e.g. the hyperlink
/// style for URLs).
pub fn cell_format(data: &CellData) -> WriteResult<Option<Format>> {
    match data {
        CellData::Date(_, formats) => {
            apply_formats(Format::new().set_num_format(DEFAULT_DATE_FORMAT), formats).map(Some)
        }
        CellData::DateTime(_, formats) => apply_formats(
            Format::new().set_num_format(DEFAULT_DATE_TIME_FORMAT),
            formats,
        )
        .map(Some),
        CellData::String(_, formats)
        | CellData::Number(_, formats)
        | CellData::Boolean(_, formats)
        | CellData::Formula(_, formats)
        | CellData::Url(_, _, formats)
        | CellData::Blank(formats)
        | CellData::RichString(_, formats) => optional_format(formats),
        CellData::Image(_) | CellData::ImagePath(_) => Ok(None),
    }
}

fn optional_format(formats: &[CellFormat]) -> WriteResult<Option<Format>> {
    if formats.is_empty() {
        Ok(None)
    } else {
        apply_formats(Format::new(), formats).map(Some)
    }
}

pub fn apply_formats(mut format: Format, formats: &[CellFormat]) -> WriteResult<Format> {
    for cell_format in formats {
        format = match cell_format {
            CellFormat::Bold => format.set_bold(),
            CellFormat::NumFormat(format_string) => format.set_num_format(format_string),
            CellFormat::Align(position) => match position {
                CellAlignPos::Center => format.set_align(FormatAlign::Center),
                CellAlignPos::Right => format.set_align(FormatAlign::Right),
                CellAlignPos::Left => format.set_align(FormatAlign::Left),
            },
            CellFormat::BgColor(color) => format.set_background_color(parse_hex_color(color)?),
            CellFormat::Pattern(pattern) => match pattern {
                CellPattern::Solid => format.set_pattern(FormatPattern::Solid),
                CellPattern::None => format.set_pattern(FormatPattern::None),
                CellPattern::Gray125 => format.set_pattern(FormatPattern::Gray125),
                CellPattern::Gray0625 => format.set_pattern(FormatPattern::Gray0625),
            },
            CellFormat::FontColor(color) => format.set_font_color(parse_hex_color(color)?),
            CellFormat::Italic => format.set_italic(),
            CellFormat::Underline(style) => match style {
                UnderlineStyle::Single => format.set_underline(FormatUnderline::Single),
                UnderlineStyle::Double => format.set_underline(FormatUnderline::Double),
                UnderlineStyle::SingleAccounting => {
                    format.set_underline(FormatUnderline::SingleAccounting)
                }
                UnderlineStyle::DoubleAccounting => {
                    format.set_underline(FormatUnderline::DoubleAccounting)
                }
            },
            CellFormat::Strikethrough => format.set_font_strikethrough(),
            CellFormat::FontSize(size) => format.set_font_size(font_size(*size)?),
            CellFormat::FontName(name) => format.set_font_name(name),
            CellFormat::Superscript => format.set_font_script(FormatScript::Superscript),
            CellFormat::Subscript => format.set_font_script(FormatScript::Subscript),
            CellFormat::Border(style) => format.set_border(border_style(style)),
            CellFormat::BorderTop(style) => format.set_border_top(border_style(style)),
            CellFormat::BorderBottom(style) => format.set_border_bottom(border_style(style)),
            CellFormat::BorderLeft(style) => format.set_border_left(border_style(style)),
            CellFormat::BorderRight(style) => format.set_border_right(border_style(style)),
            CellFormat::BorderColor(color) => format.set_border_color(parse_hex_color(color)?),
            CellFormat::BorderTopColor(color) => {
                format.set_border_top_color(parse_hex_color(color)?)
            }
            CellFormat::BorderBottomColor(color) => {
                format.set_border_bottom_color(parse_hex_color(color)?)
            }
            CellFormat::BorderLeftColor(color) => {
                format.set_border_left_color(parse_hex_color(color)?)
            }
            CellFormat::BorderRightColor(color) => {
                format.set_border_right_color(parse_hex_color(color)?)
            }
            CellFormat::TextWrap => format.set_text_wrap(),
            CellFormat::Valign(position) => match position {
                CellVAlignPos::Top => format.set_align(FormatAlign::Top),
                CellVAlignPos::Center => format.set_align(FormatAlign::VerticalCenter),
                CellVAlignPos::Bottom => format.set_align(FormatAlign::Bottom),
                CellVAlignPos::Justify => format.set_align(FormatAlign::VerticalJustify),
                CellVAlignPos::Distributed => format.set_align(FormatAlign::VerticalDistributed),
            },
            CellFormat::Rotation(angle) => format.set_rotation(rotation(*angle)?),
            CellFormat::Shrink => format.set_shrink(),
            CellFormat::Indent(level) => format.set_indent(indent(*level)?),
        };
    }

    Ok(format)
}

/// Parses `"#RRGGBB"` or `"RRGGBB"` (case-insensitive). Anything else is an
/// error rather than being ignored or misread.
pub fn parse_hex_color(color: &str) -> WriteResult<Color> {
    let hex = color.strip_prefix('#').unwrap_or(color);

    if hex.len() != 6 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(WriteError::new(format!(
            "invalid color {color:?}, expected a hex string like \"#RRGGBB\""
        )));
    }

    u32::from_str_radix(hex, 16)
        .map(Color::RGB)
        .map_err(|_| WriteError::new(format!("invalid color {color:?}")))
}

fn font_size(size: f64) -> WriteResult<f64> {
    if size <= 0.0 {
        Err(WriteError::new(format!(
            "font size {size} must be greater than 0"
        )))
    } else if size.is_finite() && size <= MAX_FONT_SIZE {
        Ok(size)
    } else {
        Err(WriteError::new(format!(
            "font size {size} is outside Excel's range 0..{MAX_FONT_SIZE}"
        )))
    }
}

fn rotation(angle: i64) -> WriteResult<i16> {
    match angle {
        -90..=90 | 270 => Ok(angle as i16),
        _ => Err(WriteError::new(format!(
            "rotation {angle} must be between -90 and 90, or 270"
        ))),
    }
}

fn indent(level: i64) -> WriteResult<u8> {
    u8::try_from(level)
        .ok()
        .filter(|level| *level <= MAX_INDENT)
        .ok_or_else(|| {
            WriteError::new(format!(
                "indent {level} is outside the range 0..{MAX_INDENT}"
            ))
        })
}

fn border_style(style: &BorderStyle) -> FormatBorder {
    match style {
        BorderStyle::Thin => FormatBorder::Thin,
        BorderStyle::Medium => FormatBorder::Medium,
        BorderStyle::Thick => FormatBorder::Thick,
        BorderStyle::Dashed => FormatBorder::Dashed,
        BorderStyle::Dotted => FormatBorder::Dotted,
        BorderStyle::Double => FormatBorder::Double,
        BorderStyle::Hair => FormatBorder::Hair,
        BorderStyle::MediumDashed => FormatBorder::MediumDashed,
        BorderStyle::DashDot => FormatBorder::DashDot,
        BorderStyle::MediumDashDot => FormatBorder::MediumDashDot,
        BorderStyle::DashDotDot => FormatBorder::DashDotDot,
        BorderStyle::MediumDashDotDot => FormatBorder::MediumDashDotDot,
        BorderStyle::SlantDashDot => FormatBorder::SlantDashDot,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_six_digit_hex_colors() {
        assert_eq!(parse_hex_color("#FF0000"), Ok(Color::RGB(0xFF0000)));
        assert_eq!(parse_hex_color("#00ff7f"), Ok(Color::RGB(0x00FF7F)));
        assert_eq!(parse_hex_color("4472C4"), Ok(Color::RGB(0x4472C4)));
    }

    #[test]
    fn rejects_other_color_strings() {
        for color in [
            "#FFF",
            "FFF",
            "red",
            "#FFFFFFFF",
            "",
            "#",
            "#GGGGGG",
            "#+12345",
        ] {
            assert!(
                parse_hex_color(color).is_err(),
                "{color:?} should be rejected"
            );
        }
    }

    #[test]
    fn checks_font_size() {
        assert_eq!(font_size(11.5), Ok(11.5));
        assert!(font_size(0.0).is_err());
        assert!(font_size(410.0).is_err());
    }

    #[test]
    fn checks_rotation_and_indent() {
        assert_eq!(rotation(-90), Ok(-90));
        assert_eq!(rotation(90), Ok(90));
        assert_eq!(rotation(270), Ok(270));
        assert!(rotation(91).is_err());
        assert!(rotation(65_536 + 45).is_err());

        assert_eq!(indent(250), Ok(250));
        assert!(indent(251).is_err());
        assert!(indent(-1).is_err());
    }

    #[test]
    fn rejects_invalid_colors_in_formats() {
        assert!(apply_formats(Format::new(), &[CellFormat::BgColor("#FFF".into())]).is_err());
        assert!(apply_formats(Format::new(), &[CellFormat::FontColor("#FF0000".into())]).is_ok());
    }
}
