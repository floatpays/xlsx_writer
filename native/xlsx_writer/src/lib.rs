//! The `XlsxWriter.RustXlsxWriter` NIF.
//!
//! - `protocol` decodes the private terms built by `lib/xlsx_writer.ex`.
//! - `workbook` builds the workbook from the decoded sheets.
//! - `worksheet` executes a sheet's instructions.
//! - `format` turns format options into rust_xlsxwriter formats.

mod error;
mod format;
mod protocol;
mod workbook;
mod worksheet;

use rustler::{Binary, Env, OwnedBinary, Term};

#[rustler::nif(schedule = "DirtyCpu")]
fn write<'a>(
    env: Env<'a>,
    sheets: Vec<Term<'a>>,
    properties: Term<'a>,
) -> Result<Binary<'a>, String> {
    let buffer = workbook::write_workbook(sheets, properties).map_err(|error| error.0)?;

    let mut binary = OwnedBinary::new(buffer.len())
        .ok_or_else(|| "could not allocate a binary for the workbook".to_string())?;
    binary.as_mut_slice().copy_from_slice(&buffer);

    Ok(binary.release(env))
}

rustler::init!("Elixir.XlsxWriter.RustXlsxWriter");
