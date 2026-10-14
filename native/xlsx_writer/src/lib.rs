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

use std::any::Any;
use std::panic::{self, AssertUnwindSafe};

use rustler::{Binary, Env, OwnedBinary, Term};

use crate::error::{WriteError, WriteResult};

#[rustler::nif(schedule = "DirtyCpu")]
fn write<'a>(
    env: Env<'a>,
    sheets: Vec<Term<'a>>,
    properties: Term<'a>,
) -> Result<Binary<'a>, String> {
    let buffer =
        catch_panics(|| workbook::write_workbook(sheets, properties)).map_err(|error| error.0)?;

    let mut binary = OwnedBinary::new(buffer.len())
        .ok_or_else(|| "could not allocate a binary for the workbook".to_string())?;
    binary.as_mut_slice().copy_from_slice(&buffer);

    Ok(binary.release(env))
}

/// rust_xlsxwriter may panic on input it does not expect. Report a panic like
/// any other error instead of letting Rustler raise an opaque
/// `:nif_panicked` error.
fn catch_panics<T>(operation: impl FnOnce() -> WriteResult<T>) -> WriteResult<T> {
    panic::catch_unwind(AssertUnwindSafe(operation)).unwrap_or_else(|payload| {
        Err(WriteError::new(format!(
            "could not write workbook: {}",
            panic_message(&*payload)
        )))
    })
}

fn panic_message(payload: &(dyn Any + Send)) -> &str {
    payload
        .downcast_ref::<&str>()
        .copied()
        .or_else(|| payload.downcast_ref::<String>().map(String::as_str))
        .unwrap_or("the writer panicked")
}

rustler::init!("Elixir.XlsxWriter.RustXlsxWriter");

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn returns_panics_as_errors() {
        let static_message: WriteResult<()> = catch_panics(|| panic!("static message"));
        assert_eq!(
            static_message,
            Err(WriteError::new("could not write workbook: static message"))
        );

        let formatted_message: WriteResult<()> = catch_panics(|| panic!("row {}", 7));
        assert_eq!(
            formatted_message,
            Err(WriteError::new("could not write workbook: row 7"))
        );

        let other_payload: WriteResult<()> = catch_panics(|| std::panic::panic_any(42_u8));
        assert_eq!(
            other_payload,
            Err(WriteError::new(
                "could not write workbook: the writer panicked"
            ))
        );
    }

    #[test]
    fn passes_results_through() {
        assert_eq!(catch_panics(|| Ok(1)), Ok(1));
        assert_eq!(
            catch_panics::<()>(|| Err(WriteError::new("bad"))),
            Err(WriteError::new("bad"))
        );
    }
}
