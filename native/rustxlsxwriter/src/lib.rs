use rust_xlsxwriter::{ExcelDateTime, Format, FormatAlign, Image, Workbook};
use rustler::{Binary, NifTaggedEnum};

#[derive(NifTaggedEnum, PartialEq)]
enum CellAlignPos {
    Center,
    Left,
    Right
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
enum Instruction<'a> {
    Write(u32, u16, CellData<'a>),
    SetColumnWidth(u16, u32),
    SetRowHeight(u32, u16),
}

#[rustler::nif]
fn write(instructions: Vec<Instruction>) -> Result<Vec<u8>, String> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    for instruction in instructions {
        let _result = match instruction {
            Instruction::SetColumnWidth(col, val) => worksheet.set_column_width(col, val),
            Instruction::SetRowHeight(row, val) => worksheet.set_row_height(row, val),
            Instruction::Write(col, row, data) => match data {
                CellData::String(val) => worksheet.write(col, row, val),
                CellData::StringWithFormat(val, formats) => {
                    let mut merge_format = Format::new();

                    for format in formats {
                        merge_format = match format {
                            CellFormat::Bold => merge_format.set_bold(),
                            CellFormat::Align(pos) => match pos {
                                CellAlignPos::Center => merge_format.set_align(FormatAlign::Center),
                                CellAlignPos::Right => merge_format.set_align(FormatAlign::Right),
                                CellAlignPos::Left => merge_format.set_align(FormatAlign::Left),
                            },
                        }
                    }

                    worksheet.write_with_format(col, row, val, &merge_format)
                }
                CellData::Float(val) => worksheet.write(col, row, val),
                CellData::Date(year, month, day) => {
                    let date_format = Format::new().set_num_format("yyyy-mm-dd");

                    match ExcelDateTime::from_ymd(year, month, day) {
                        Err(e) => return Err(e.to_string()),
                        Ok(date) => worksheet.write_with_format(6, 0, &date, &date_format),
                    }
                }
                CellData::ImagePath(val) => match Image::new(val) {
                    Err(e) => return Err(e.to_string()),
                    Ok(image) => worksheet.insert_image(col, row, &image),
                },
                CellData::Image(binary) => {
                    let val = binary.as_slice().to_vec();

                    match Image::new_from_buffer(&val) {
                        Err(e) => return Err(e.to_string()),
                        Ok(image) => worksheet.insert_image(col, row, &image),
                    }
                }
            },
        };
    }

    match workbook.save_to_buffer() {
        Ok(buf) => return Ok(buf),
        Err(e) => {
            // Return an atom saying there was an error.
            // We can figure out later how to include more data
            // about the error.
            return Err(e.to_string());
        }
    }
}

rustler::init!("Elixir.XlsxWriter.RustXlsxWriter");
