use rust_xlsxwriter::{Image, Workbook};
use rustler::{Binary, NifTaggedEnum};

#[derive(NifTaggedEnum)]
enum CellData<'a> {
    Float(f64),
    String(String),
    ImagePath(String),
    Image(Binary<'a>),
}

#[derive(NifTaggedEnum)]
enum Instruction<'a> {
    Insert(u32, u16, CellData<'a>),
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
            Instruction::Insert(col, row, data) => match data {
                CellData::String(val) => worksheet.write_string(col, row, val),
                CellData::Float(val) => worksheet.write_number(col, row, val),
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
