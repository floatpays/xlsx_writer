use rust_xlsxwriter::{Color, ExcelDateTime, Format, FormatAlign, FormatBorder, FormatPattern, FormatScript, FormatUnderline, Image, Workbook, Worksheet, XlsxError, Formula, Url};
use rustler::{Binary, NifTaggedEnum};

#[derive(NifTaggedEnum, PartialEq)]
enum CellAlignPos {
    Center,
    Left,
    Right,
}

#[derive(NifTaggedEnum, PartialEq)]
enum CellPattern {
    Solid,
    None,
    Gray125,
    Gray0625,
}

#[derive(NifTaggedEnum, PartialEq)]
enum UnderlineStyle {
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

#[derive(NifTaggedEnum, PartialEq)]
enum BorderStyle {
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

#[derive(NifTaggedEnum, PartialEq)]
enum CellFormat {
    Bold,
    Align(CellAlignPos),
    // Examples of numeric formats: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Format.html#examples-2
    NumFormat(String),
    BgColor(String),
    Pattern(CellPattern),
    FontColor(String),
    Italic,
    Underline(UnderlineStyle),
    Strikethrough,
    FontSize(u16),
    FontName(String),
    Superscript,
    Subscript,
    Border(BorderStyle),
    BorderTop(BorderStyle),
    BorderBottom(BorderStyle),
    BorderLeft(BorderStyle),
    BorderRight(BorderStyle),
    BorderColor(String),
    BorderTopColor(String),
    BorderBottomColor(String),
    BorderLeftColor(String),
    BorderRightColor(String),
}

#[derive(NifTaggedEnum)]
enum CellData<'a> {
    Float(f64),
    String(String),
    StringWithFormat(String, Vec<CellFormat>),
    NumberWithFormat(f64, Vec<CellFormat>),
    ImagePath(String),
    Image(Binary<'a>),
    Date(String),
    DateTime(String),
    Formula(String),
    Boolean(bool),
    BooleanWithFormat(bool, Vec<CellFormat>),
    Url(String),
    UrlWithText(String, String),
    UrlWithFormat(String, Vec<CellFormat>),
    UrlWithTextAndFormat(String, String, Vec<CellFormat>),
    Blank(Vec<CellFormat>),
}

#[derive(NifTaggedEnum)]
enum Sheet<'a> {
    Write(u32, u16, CellData<'a>),
    SetColumnWidth(u16, u32),
    SetRowHeight(u32, u16),
    SetFreezePanes(u32, u16),
    SetRowHidden(u32),
    SetColumnHidden(u16),
    SetAutofilter(u32, u16, u32, u16),
    MergeRange(u32, u16, u32, u16, CellData<'a>),
}

#[rustler::nif]
fn write(sheets: Vec<(String, Vec<Sheet>)>) -> Result<Vec<u8>, String> {
    let mut workbook = Workbook::new();

    for (sheet_name, sheet) in sheets {
        let worksheet = workbook.add_worksheet();

        match worksheet.set_name(sheet_name) {
            Err(e) => return Err(e.to_string()),
            Ok(_) => (),
        }

        for instruction in sheet {
            let _result = match instruction {
                Sheet::SetColumnWidth(col, val) => worksheet.set_column_width(col, val),
                Sheet::SetRowHeight(row, val) => worksheet.set_row_height(row, val),
                Sheet::SetFreezePanes(row, col) => worksheet.set_freeze_panes(row, col),
                Sheet::SetRowHidden(row) => worksheet.set_row_hidden(row),
                Sheet::SetColumnHidden(col) => worksheet.set_column_hidden(col),
                Sheet::SetAutofilter(first_row, first_col, last_row, last_col) => {
                    worksheet.autofilter(first_row, first_col, last_row, last_col)
                }
                Sheet::MergeRange(first_row, first_col, last_row, last_col, data) => {
                    merge_range(worksheet, first_row, first_col, last_row, last_col, data)
                }
                Sheet::Write(row, col, data) => write_data(worksheet, row, col, data),
            };
        }
    }

    return match workbook.save_to_buffer() {
        Ok(buf) => Ok(buf),
        Err(e) => Err(e.to_string()),
    };
}

fn merge_range<'a, 'b>(
    worksheet: &'a mut Worksheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    data: CellData<'b>,
) -> Result<&'a mut Worksheet, XlsxError> {
    match data {
        CellData::String(val) => worksheet.merge_range(first_row, first_col, last_row, last_col, &val, &Format::new()),
        CellData::StringWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.merge_range(first_row, first_col, last_row, last_col, &val, &format)
        }
        CellData::NumberWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            // Write value to first cell, then merge the range with the same format
            worksheet.write_number_with_format(first_row, first_col, val, &format)?;
            worksheet.merge_range(first_row, first_col, last_row, last_col, "", &format)
        }
        CellData::Float(val) => {
            // Write number to first cell, then merge
            worksheet.write_number(first_row, first_col, val)?;
            worksheet.merge_range(first_row, first_col, last_row, last_col, "", &Format::new())
        }
        CellData::Boolean(val) => {
            // Write boolean to first cell, then merge
            worksheet.write_boolean(first_row, first_col, val)?;
            worksheet.merge_range(first_row, first_col, last_row, last_col, "", &Format::new())
        }
        CellData::BooleanWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.write_boolean_with_format(first_row, first_col, val, &format)?;
            worksheet.merge_range(first_row, first_col, last_row, last_col, "", &format)
        }
        CellData::Blank(formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.merge_range(first_row, first_col, last_row, last_col, "", &format)
        }
        // For other types that don't support merge_range, write to first cell only
        _ => write_data(worksheet, first_row, first_col, data),
    }
}

fn write_data<'a, 'b>(
    worksheet: &'a mut Worksheet,
    row: u32,
    col: u16,
    data: CellData<'b>,
) -> Result<&'a mut Worksheet, XlsxError> {
    match data {
        CellData::String(val) => worksheet.write(row, col, val),
        CellData::StringWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.write_with_format(row, col, val, &format)
        }
        CellData::NumberWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.write_number_with_format(row, col, val, &format)
        }

        CellData::Float(val) => worksheet.write(row, col, val),
        CellData::Date(iso8601) => {
            let date_format = Format::new().set_num_format("yyyy-mm-dd");

            match ExcelDateTime::parse_from_str(&iso8601) {
                Err(e) => return Err(e),
                Ok(date) => worksheet.write_with_format(row, col, &date, &date_format),
            }
        },
        CellData::DateTime(iso8601) => {
            let date_format = Format::new().set_num_format("yyyy-mm-ddThh:mm:ss");

            match ExcelDateTime::parse_from_str(&iso8601) {
                Err(e) => return Err(e),
                Ok(date) => worksheet.write_with_format(row, col, &date, &date_format),
            }
        },
        CellData::Formula(val) => worksheet.write(row, col, Formula::new(val)),
        CellData::Boolean(val) => worksheet.write_boolean(row, col, val),
        CellData::BooleanWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.write_boolean_with_format(row, col, val, &format)
        }
        CellData::Url(url) => {
            let url_obj = Url::new(&url);
            worksheet.write_url(row, col, &url_obj)
        }
        CellData::UrlWithText(url, text) => {
            let url_obj = Url::new(&url);
            worksheet.write_url_with_text(row, col, &url_obj, &text)
        }
        CellData::UrlWithFormat(url, formats) => {
            let format = apply_formats(Format::new(), &formats);
            let url_obj = Url::new(&url);
            worksheet.write_url_with_format(row, col, &url_obj, &format)
        }
        CellData::UrlWithTextAndFormat(url, text, formats) => {
            let format = apply_formats(Format::new(), &formats);
            let url_obj = Url::new(&url);
            worksheet.write_url_with_text(row, col, &url_obj, &text)?;
            worksheet.write_with_format(row, col, &text, &format)
        }
        CellData::Blank(formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.write_blank(row, col, &format)
        }
        CellData::ImagePath(val) => match Image::new(val) {
            Err(e) => return Err(e),
            Ok(image) => worksheet.insert_image(row, col, &image),
        },
        CellData::Image(binary) => {
            let val = binary.as_slice().to_vec();

            match Image::new_from_buffer(&val) {
                Err(e) => return Err(e),
                Ok(image) => worksheet.insert_image(row, col, &image),
            }
        }
    }
}

fn apply_formats(mut format: Format, formats: &[CellFormat]) -> Format {
    for fmt in formats {
        format = match fmt {
            CellFormat::Bold => format.set_bold(),
            CellFormat::NumFormat(format_string) => format.set_num_format(format_string),
            CellFormat::Align(pos) => match pos {
                CellAlignPos::Center => format.set_align(FormatAlign::Center),
                CellAlignPos::Right => format.set_align(FormatAlign::Right),
                CellAlignPos::Left => format.set_align(FormatAlign::Left),
            },
            CellFormat::BgColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_background_color(color)
                } else {
                    format
                }
            }
            CellFormat::Pattern(pattern) => match pattern {
                CellPattern::Solid => format.set_pattern(FormatPattern::Solid),
                CellPattern::None => format.set_pattern(FormatPattern::None),
                CellPattern::Gray125 => format.set_pattern(FormatPattern::Gray125),
                CellPattern::Gray0625 => format.set_pattern(FormatPattern::Gray0625),
            },
            CellFormat::FontColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_font_color(color)
                } else {
                    format
                }
            }
            CellFormat::Italic => format.set_italic(),
            CellFormat::Underline(style) => match style {
                UnderlineStyle::Single => format.set_underline(FormatUnderline::Single),
                UnderlineStyle::Double => format.set_underline(FormatUnderline::Double),
                UnderlineStyle::SingleAccounting => format.set_underline(FormatUnderline::SingleAccounting),
                UnderlineStyle::DoubleAccounting => format.set_underline(FormatUnderline::DoubleAccounting),
            },
            CellFormat::Strikethrough => format.set_font_strikethrough(),
            CellFormat::FontSize(size) => format.set_font_size(*size),
            CellFormat::FontName(name) => format.set_font_name(name),
            CellFormat::Superscript => format.set_font_script(FormatScript::Superscript),
            CellFormat::Subscript => format.set_font_script(FormatScript::Subscript),
            CellFormat::Border(style) => {
                let border_style = convert_border_style(style);
                format.set_border(border_style)
            },
            CellFormat::BorderTop(style) => {
                let border_style = convert_border_style(style);
                format.set_border_top(border_style)
            },
            CellFormat::BorderBottom(style) => {
                let border_style = convert_border_style(style);
                format.set_border_bottom(border_style)
            },
            CellFormat::BorderLeft(style) => {
                let border_style = convert_border_style(style);
                format.set_border_left(border_style)
            },
            CellFormat::BorderRight(style) => {
                let border_style = convert_border_style(style);
                format.set_border_right(border_style)
            },
            CellFormat::BorderColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_border_color(color)
                } else {
                    format
                }
            },
            CellFormat::BorderTopColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_border_top_color(color)
                } else {
                    format
                }
            },
            CellFormat::BorderBottomColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_border_bottom_color(color)
                } else {
                    format
                }
            },
            CellFormat::BorderLeftColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_border_left_color(color)
                } else {
                    format
                }
            },
            CellFormat::BorderRightColor(color_hex) => {
                if let Some(color) = parse_hex_color(color_hex) {
                    format.set_border_right_color(color)
                } else {
                    format
                }
            },
        };
    }
    return format;
}

/// Parses a hex color string (e.g., "#FF0000" or "FF0000") into a Color.
/// Returns None if the hex string is invalid.
fn parse_hex_color(color_hex: &str) -> Option<Color> {
    let hex_str = color_hex.trim_start_matches('#');
    u32::from_str_radix(hex_str, 16)
        .ok()
        .map(Color::from)
}

fn convert_border_style(style: &BorderStyle) -> FormatBorder {
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

rustler::init!("Elixir.XlsxWriter.RustXlsxWriter");
