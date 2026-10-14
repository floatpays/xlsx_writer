# xlsx_writer NIF

The Rust side of `XlsxWriter`, loaded as `XlsxWriter.RustXlsxWriter` (see
`lib/xlsx_writer/rust_xlsx_writer.ex`). It builds the workbook with
[rust_xlsxwriter](https://docs.rs/rust_xlsxwriter) and returns it as a binary.

- `src/lib.rs` - the `write/2` NIF, which turns panics into error tuples.
- `src/protocol.rs` - decoding the sheets and instructions built by
  `lib/xlsx_writer.ex`. These terms are private to the library.
- `src/workbook.rs` - building the workbook and its document properties.
- `src/worksheet.rs` - executing a sheet's instructions (cells, merges,
  sizes, comments, autofit).
- `src/format.rs` - turning format options into rust_xlsxwriter formats.
- `src/error.rs` - the error type and short descriptions of invalid terms.

Excel's limits are checked in Elixir by `XlsxWriter.Validation`; the NIF only
checks what it needs to represent a value (integer narrowing, finite floats,
colours).

Precompiled binaries are built by `.github/workflows/release.yml` for NIF 2.16;
set `config :rustler_precompiled, :force_build, xlsx_writer: true` to build
from source (this repository's `config/config.exs` does so for development).

Run `cargo fmt`, `cargo clippy --all-targets -- -D warnings` and the crate's
unit tests from this directory.
