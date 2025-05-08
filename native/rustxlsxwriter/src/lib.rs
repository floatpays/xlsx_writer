use rust_xlsxwriter::{workbook, ExcelDateTime, Format, Image, Workbook, XlsxError};
use rustler::Atom;

mod atoms {
    rustler::atoms! {
        xlsx_generation_error,
    }
}

// We can't write an Encoder for XlsxError, due to rust's orphan
// rule. We can only implement traits for types we own, or
// types that are defined in the same crate as the trait. In our case
// both the trait(Encoder) and the type (XlsxError) are defined in different crates.
//
// impl Encoder for XlsxError
// {
//     fn encode<'c>(&self, env: Env<'c>) -> Term<'c> {
//         ...return encoded version of XlsxError
//     }
// }

fn add_worksheet_examples(workbook: &mut Workbook) -> Result<(), XlsxError> {
    // Scenarios we need to support:
    // 1. Add a sheet with a name
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("custom_sheet_name")?;

    // 2. Add text to a cell.
    worksheet.write(0, 0, "We can write plain text")?;

    // 3. Add a formatted numeric value to a cell:
    worksheet.write(1, 1, "<- We can write formatted numbers")?;
    let decimal_format = Format::new().set_num_format("0.000");
    worksheet.write_with_format(1, 0, 3.50, &decimal_format)?;

    // 4. Add formatted Date to a cell:
    worksheet.write(2, 1, "<- We can write dates")?;
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let date = ExcelDateTime::from_ymd(2023, 1, 25)?;
    worksheet.write_with_format(2, 0, &date, &date_format)?;

    // 5. Add an image to a cell:
    let image = Image::new("bird.jpeg")?;
    worksheet.insert_image(3, 0, &image)?;

    // 6. Return worksheet as a binary, don't save to disk.
    // TODO...

    // 7. Pass in image as a binary, not a filesystem location.
    // (maybe filesystem location is OK? If we have the standard bank logo in our repo? TBD...)
    // TODO...

    Ok(())
}

#[rustler::nif]
fn add(a: i64, b: i64) -> i64 {
    a + b
}

// Return some binary data. We'll need this to get the worksheet.
// A byte vector is what Workbook.save_to_buffer returns.
#[rustler::nif]
fn get_binary() -> Result<Vec<u8>, Atom> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    if let Err(_xlsx_error) = worksheet.write_string(0, 0, "Hello this is from binary") {
        // TODO: Log the actual error details somewhere. Or return it somehow.
        return Err(atoms::xlsx_generation_error());
    }

    match workbook.save_to_buffer() {
        Ok(buf) => return Ok(buf),
        Err(_e) => {
            // Return an atom saying there was an error.
            // We can figure out later how to include more data
            // about the error.
            return Err(atoms::xlsx_generation_error());
        }
    }
}

// Given an image as a binary (not file path), return a binary worksheet.
#[rustler::nif]
fn get_binary_with_image(image_bytes_vector: Vec<u8>) -> Result<Vec<u8>, Atom> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let image_bytes: &[u8] = &image_bytes_vector;

    // TODO: Is there something like a 'with' statement in Rust?
    // (Using "if let" maybe?)
    match Image::new_from_buffer(image_bytes) {
        Ok(image) => {
            match worksheet.insert_image(3, 0, &image) {
                Ok(_) => {
                    //                return Ok(image_bytes_vector);
                    match workbook.save_to_buffer() {
                        Ok(buf) => return Ok(buf),
                        Err(_e) => {
                            // Something went wrong when savings workbook to buffer.
                            return Err(atoms::xlsx_generation_error());
                        }
                    }
                }
                Err(_xlsx_error) => {
                    // Something went wrong when inserting the image.
                    return Err(atoms::xlsx_generation_error());
                }
            }
        }
        Err(_e) => {
            // Something went wrong when creating the image from buffer.
            return Err(atoms::xlsx_generation_error());
        }
    }
}

#[rustler::nif]
fn test_new_workbook() -> Result<(), Atom> {
    let mut workbook = Workbook::new();

    if let Err(_xlsx_error) = add_worksheet_examples(&mut workbook) {
        // TODO: Log the actual error details somewhere. Or return it somehow.
        return Err(atoms::xlsx_generation_error());
    }

    match workbook.save("demo.xlsx") {
        Ok(_) => return Ok(()),
        Err(_e) => {
            // Return an atom saying there was an error.
            // We can figure out later how to include more data
            // about the error.
            return Err(atoms::xlsx_generation_error());
        }
    }
}

rustler::init!("Elixir.XlsxWriter");
