use rust_xlsxwriter::{Color, ExcelDateTime, Format, FormatAlign, FormatBorder, FormatPattern, FormatScript, FormatUnderline, Image, Note, Workbook, Worksheet, XlsxError, Formula, Url};
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

#[derive(rustler::NifStruct)]
#[module = "XlsxWriter.NoteOptions"]
struct NoteOptions {
    author: Option<String>,
    visible: Option<bool>,
    width: Option<u32>,
    height: Option<u32>,
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
    RichString(Vec<(String, Vec<CellFormat>)>),
    RichStringWithFormat(Vec<(String, Vec<CellFormat>)>, Vec<CellFormat>),
}

#[derive(NifTaggedEnum)]
enum Sheet<'a> {
    Write(u32, u16, CellData<'a>),
    SetColumnWidth(u16, u32),
    SetRowHeight(u32, u16),
    SetColumnRangeWidth(u16, u16, u32),
    SetRowRangeHeight(u32, u32, u16),
    SetFreezePanes(u32, u16),
    SetRowHidden(u32),
    SetColumnHidden(u16),
    SetAutofilter(u32, u16, u32, u16),
    MergeRange(u32, u16, u32, u16, CellData<'a>),
    InsertNote(u32, u16, String, NoteOptions),
}

#[rustler::nif]
fn write(sheets: Vec<(String, Vec<Sheet>)>) -> Result<Vec<u8>, String> {
    let mut workbook = Workbook::new();

    for (sheet_name, sheet) in sheets {
        let mut worksheet = workbook.add_worksheet();

        match worksheet.set_name(sheet_name) {
            Err(e) => return Err(e.to_string()),
            Ok(_) => (),
        }

        for instruction in sheet {
            worksheet = match instruction {
                Sheet::SetColumnWidth(col, val) => match worksheet.set_column_width(col, val) {
                    Ok(ws) => ws,
                    Err(e) => return Err(e.to_string()),
                },
                Sheet::SetRowHeight(row, val) => match worksheet.set_row_height(row, val) {
                    Ok(ws) => ws,
                    Err(e) => return Err(e.to_string()),
                },
                Sheet::SetColumnRangeWidth(first_col, last_col, width) => {
                    // Set width for each column in the range
                    let mut ws = worksheet;
                    for col in first_col..=last_col {
                        ws = match ws.set_column_width_pixels(col, width) {
                            Ok(w) => w,
                            Err(e) => return Err(e.to_string()),
                        };
                    }
                    ws
                }
                Sheet::SetRowRangeHeight(first_row, last_row, height) => {
                    // Set height for each row in the range
                    let mut ws = worksheet;
                    for row in first_row..=last_row {
                        ws = match ws.set_row_height_pixels(row, height as u32) {
                            Ok(w) => w,
                            Err(e) => return Err(e.to_string()),
                        };
                    }
                    ws
                }
                Sheet::SetFreezePanes(row, col) => match worksheet.set_freeze_panes(row, col) {
                    Ok(ws) => ws,
                    Err(e) => return Err(e.to_string()),
                },
                Sheet::SetRowHidden(row) => match worksheet.set_row_hidden(row) {
                    Ok(ws) => ws,
                    Err(e) => return Err(e.to_string()),
                },
                Sheet::SetColumnHidden(col) => match worksheet.set_column_hidden(col) {
                    Ok(ws) => ws,
                    Err(e) => return Err(e.to_string()),
                },
                Sheet::SetAutofilter(first_row, first_col, last_row, last_col) => {
                    match worksheet.autofilter(first_row, first_col, last_row, last_col) {
                        Ok(ws) => ws,
                        Err(e) => return Err(e.to_string()),
                    }
                }
                Sheet::MergeRange(first_row, first_col, last_row, last_col, data) => {
                    match merge_range(worksheet, first_row, first_col, last_row, last_col, data) {
                        Ok(ws) => ws,
                        Err(e) => return Err(e.to_string()),
                    }
                }
                Sheet::InsertNote(row, col, text, options) => {
                    match insert_note(worksheet, row, col, text, options) {
                        Ok(ws) => ws,
                        Err(e) => return Err(e.to_string()),
                    }
                }
                Sheet::Write(row, col, data) => match write_data(worksheet, row, col, data) {
                    Ok(ws) => ws,
                    Err(e) => return Err(e.to_string()),
                },
            };
        }
    }

    return match workbook.save_to_buffer() {
        Ok(buf) => Ok(buf),
        Err(e) => Err(e.to_string()),
    };
}

fn insert_note<'a>(
    worksheet: &'a mut Worksheet,
    row: u32,
    col: u16,
    text: String,
    options: NoteOptions,
) -> Result<&'a mut Worksheet, XlsxError> {
    let mut note = Note::new(&text);

    if let Some(author) = options.author {
        note = note.set_author(&author);
    }

    if let Some(visible) = options.visible {
        note = note.set_visible(visible);
    }

    if let Some(width) = options.width {
        note = note.set_width(width);
    }

    if let Some(height) = options.height {
        note = note.set_height(height);
    }

    worksheet.insert_note(row, col, &note)
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
        CellData::RichString(segments) => {
            write_rich_string_helper(worksheet, row, col, &segments, None)
        }
        CellData::RichStringWithFormat(segments, cell_formats) => {
            let cell_format = apply_formats(Format::new(), &cell_formats);
            write_rich_string_helper(worksheet, row, col, &segments, Some(cell_format))
        }
    }
}

fn write_rich_string_helper<'a>(
    worksheet: &'a mut Worksheet,
    row: u32,
    col: u16,
    segments: &[(String, Vec<CellFormat>)],
    cell_format: Option<Format>,
) -> Result<&'a mut Worksheet, XlsxError> {
    // Build format objects for each segment
    let segment_formats: Vec<Format> = segments
        .iter()
        .map(|(_, formats)| apply_formats(Format::new(), formats))
        .collect();

    // Build the segments array with references
    let rich_segments: Vec<(&Format, &str)> = segments
        .iter()
        .zip(segment_formats.iter())
        .map(|((text, _), format)| (format, text.as_str()))
        .collect();

    match cell_format {
        Some(format) => worksheet.write_rich_string_with_format(row, col, &rich_segments, &format),
        None => worksheet.write_rich_string(row, col, &rich_segments),
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
