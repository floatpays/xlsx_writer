use rust_xlsxwriter::{Image, Workbook};
use rustler::NifTaggedEnum;
use std::fmt;

#[derive(NifTaggedEnum)]
enum CellData {
    Float(f64),
    String(String),
    ImagePath(String),
    Image(Vec<u8>),
}

impl fmt::Display for CellData {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            CellData::Float(val) => write!(f, "Float: {}", val),
            CellData::String(val) => write!(f, "String: {}", val),
            CellData::ImagePath(val) => write!(f, "Image: {}", val),
            CellData::Image(_val) => write!(f, "Image: <<binary>>"),
        }
    }
}

#[rustler::nif]
fn write(data: Vec<(u32, u16, CellData)>) -> Result<Vec<u8>, String> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    for (row, col, data) in data {
        let _ = match data {
            CellData::String(val) => worksheet.write_string(row, col, val),
            CellData::Float(val) => worksheet.write_number(row, col, val),
            CellData::ImagePath(val) => match Image::new(val) {
                Err(e) => return Err(e.to_string()),

                Ok(image) => match worksheet.insert_image(row, col, &image) {
                    Ok(val) => Ok(val),
                    Err(e) => return Err(e.to_string()),
                },
            },
            CellData::Image(val) => match Image::new_from_buffer(&val) {
                Err(e) => return Err(e.to_string()),

                Ok(image) => match worksheet.insert_image(row, col, &image) {
                    Ok(val) => Ok(val),
                    Err(e) => return Err(e.to_string()),
                },
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
