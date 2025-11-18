use rust_xlsxwriter::{ExcelDateTime, Format, FormatAlign, Image, Workbook, Worksheet, XlsxError, Formula, Url};
use rustler::{Binary, NifTaggedEnum};

#[derive(NifTaggedEnum, PartialEq)]
enum CellAlignPos {
    Center,
    Left,
    Right,
}

#[derive(NifTaggedEnum, PartialEq)]
enum CellFormat {
    Bold,
    Align(CellAlignPos),
    // Examples of numeric formats: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Format.html#examples-2
    NumFormat(String)
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
        };
    }
    return format;
}

rustler::init!("Elixir.XlsxWriter.RustXlsxWriter");
