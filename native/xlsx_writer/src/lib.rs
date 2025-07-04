use rust_xlsxwriter::{ExcelDateTime, Format, FormatAlign, Image, Workbook, Worksheet, XlsxError, Formula};
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
}

#[derive(NifTaggedEnum)]
enum Sheet<'a> {
    Write(u32, u16, CellData<'a>),
    SetColumnWidth(u16, u32),
    SetRowHeight(u32, u16),
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
                Sheet::Write(row, col, data) => write_data(worksheet, row, col, data),
            };
        }
    }

    return match workbook.save_to_buffer() {
        Ok(buf) => Ok(buf),
        Err(e) => Err(e.to_string()),
    };
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
