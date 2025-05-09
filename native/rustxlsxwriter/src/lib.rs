use rust_xlsxwriter::{ExcelDateTime, Format, FormatAlign, Image, Workbook, Worksheet, XlsxError};
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
}

#[derive(NifTaggedEnum)]
enum CellData<'a> {
    Float(f64),
    String(String),
    StringWithFormat(String, Vec<CellFormat>),
    ImagePath(String),
    Image(Binary<'a>),
    Date(u16, u8, u8),
}

#[derive(NifTaggedEnum)]
enum Sheet<'a> {
    Write(u32, u16, CellData<'a>),
    SetColumnWidth(u16, u32),
    SetRowHeight(u32, u16),
}

#[rustler::nif]
fn write(sheets: Vec<Vec<Sheet>>) -> Result<Vec<u8>, String> {
    let mut workbook = Workbook::new();

    for sheet in sheets {
        let worksheet = workbook.add_worksheet();

        for instruction in sheet {
            let _result = match instruction {
                Sheet::SetColumnWidth(col, val) => worksheet.set_column_width(col, val),
                Sheet::SetRowHeight(row, val) => worksheet.set_row_height(row, val),
                Sheet::Write(col, row, data) => write_data(worksheet, col, row, data),
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
    col: u32,
    row: u16,
    data: CellData<'b>,
) -> Result<&'a mut Worksheet, XlsxError> {
    match data {
        CellData::String(val) => worksheet.write(col, row, val),
        CellData::StringWithFormat(val, formats) => {
            let format = apply_formats(Format::new(), &formats);
            worksheet.write_with_format(col, row, val, &format)
        }
        CellData::Float(val) => worksheet.write(col, row, val),
        CellData::Date(year, month, day) => {
            let date_format = Format::new().set_num_format("yyyy-mm-dd");

            match ExcelDateTime::from_ymd(year, month, day) {
                Err(e) => return Err(e),
                Ok(date) => worksheet.write_with_format(6, 0, &date, &date_format),
            }
        }
        CellData::ImagePath(val) => match Image::new(val) {
            Err(e) => return Err(e),
            Ok(image) => worksheet.insert_image(col, row, &image),
        },
        CellData::Image(binary) => {
            let val = binary.as_slice().to_vec();

            match Image::new_from_buffer(&val) {
                Err(e) => return Err(e),
                Ok(image) => worksheet.insert_image(col, row, &image),
            }
        }
    }
}

fn apply_formats(mut format: Format, formats: &[CellFormat]) -> Format {
    for fmt in formats {
        format = match fmt {
            CellFormat::Bold => format.set_bold(),
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
