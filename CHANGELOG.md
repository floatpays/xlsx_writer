# Changelog

## v0.10.4

### bug fixes

- `write_formula/5` with `format:` now applies the format when the formula uses a dynamic array function (e.g. `FILTER`, `UNIQUE`, `SORT`). Previously the format was silently dropped (`rust_xlsxwriter` 0.99.1, [issue #190](https://github.com/jmcnamara/rust_xlsxwriter/issues/190)).

### improvements

- Update `rust_xlsxwriter` from 0.99.0 to 0.99.1. Besides the fix above, the release is internal clippy refactors with no behaviour change.
  - Verified that generated workbooks are unchanged: every example in `examples/` was generated against both the old and new dependencies, and the unzipped workbook XML was byte-identical. The only differences were the creation timestamp in `docProps/core.xml` and the random cell values in `large_dataset`.
- Update cargo transitive dependencies (cfg-if, crc32fast, equivalent, flate2, indexmap, inventory, libc, unicode-ident, zlib-rs, zopfli). `miniz_oxide` and `adler2` drop out of the lockfile.
- flate2 1.1.10 moved zlib-rs runtime CPU feature detection (SIMD) behind a new default feature, which `zip` disables. `flate2` is now listed directly in `Cargo.toml` with `runtime_detection` enabled, so compression keeps using SIMD. The backend is still zlib-rs.
- `rustler` stays on 0.38.0: crate 0.39.0 is out, but the hex package is still 0.38.0 and the two must match.
- Tested against `rustler_precompiled` 0.10.0 (allowed by the existing `~> 0.8` constraint; adds NIF 2.18 / OTP 29 support, with no change to the options or checksum format this library uses) and `decimal` 3.2.0 (fixes `Decimal.to_float/1` for odd integers between 2^52 and 2^53, which `write/4` uses for `Decimal` values).
- Update mix dev-only dependencies (`ex_doc` 0.40.4, `igniter` 0.8.4) and their transitives (ex_ast, finch, hpax, makeup, mint, req, sourceror, spitfire). Drop the stale `toml` lock entry, which no environment resolved.
- Remove the unused `dialyxir` dev/test dependency.
- CI: update actions to their current major versions.

## v0.10.3

### improvements

- Update `rust_xlsxwriter` from 0.97.0 to 0.99.0
  - 0.97.1 fixes an XML element ordering bug that produced invalid Excel files when `set_tab_color()` and `set_print_fit_to_pages()` were both applied to a worksheet. Only `set_tab_color/2` is exposed here, so this library could not trigger it.
  - 0.98.0 only bumps the optional `polars` dependency, which this library does not enable.
  - 0.99.0 renames `XlsxError::TableNameReused` to `XlsxError::NameReused` and tightens validation of table and defined names. The rename is a breaking change in the crate, but this library never matches on `XlsxError` variants (it formats them with `to_string()`) and exposes neither tables nor defined names, so there is no impact on callers.
  - Verified that generated workbooks are unchanged: running the full feature demo and an image workbook against both 0.97.0 and 0.99.0 produced byte-identical worksheet XML, styles, shared strings, drawings, notes, relationships, and embedded image bytes. The only difference was the creation timestamp in `docProps/core.xml`.
- Update cargo transitive dependencies (crc32fast, log, zlib-rs).
- Update mix dev-only dependencies (`req` 0.7.1 to 0.7.3, `spitfire` 0.3.13 to 0.4.0) and drop the stale `castore` lock entry, which no environment resolved.

### notes

- Building the NIF from source still requires Rust 1.88.0 or later; `rust_xlsxwriter` 0.99.0 did not raise its MSRV.

## v0.10.2

### improvements

- Update `rust_xlsxwriter` from 0.96.0 to 0.97.0
  - Picks up an update of the underlying `zip` crate from 7.2 to 8.6, which fixes two upstream ZIP64 "large file" handling bugs. Verified that generated workbooks are unchanged in structure: cell values, styles, hyperlinks, shared strings, autofit widths, and embedded image bytes all round-trip identically.
  - Also fixes BMP images using "top-down" DIB row order, and returns an error instead of panicking on invalid or truncated image data — both reachable through `write_image/4`.
  - No API changes affecting this library.
- Update cargo transitive dependencies (libc, proc-macro2, quote).
- Update mix dev-only dependencies (`igniter` 0.8.2 to 0.8.3) and their transitives (earmark_parser, ex_ast, glob_ex, req).

### notes

- Building the NIF from source now requires Rust 1.88.0 or later (raised by `rust_xlsxwriter` 0.97.0). This does not affect users of the precompiled binaries, which cover the same targets and NIF version as before.

## v0.10.1

### improvements

- Update `rust_xlsxwriter` from 0.95.0 to 0.96.0
  - `Workbook::save_to_writer()` no longer requires the `Seek` trait (internal move to the zip.rs stream API); no API changes affecting this library, which uses `save_to_buffer()`.
- Update cargo transitive dependencies (memchr, rustversion, simd-adler32, syn, zlib-rs).
- Update mix dev/test transitive dependencies (erlex, ex_ast, hpax, makeup, mint, req).
- CI: test against all supported Elixir versions — add Elixir 1.20 and OTP 28.5/29.0 to the matrix, and bump existing entries to the latest patch releases (Elixir 1.18.4, 1.19.5, 1.20.2 on OTP 26.2/27.3/28.5/29.0 per the official compatibility table).

## v0.10.0

### improvements

- Update `rustler` from 0.37.4 to 0.38.0 (Elixir and Rust crate).
- Update dev/test tooling: `credo`, `ex_doc`, `igniter`, and `quokka`.
- Update `decimal` from 3.1.0 to 3.1.1.
- Update cargo transitive dependencies (bumpalo, log, memchr, quote, syn, zlib-rs).

## v0.9.0

### breaking changes

- `XlsxWriter.set_column_range_width/4` and `XlsxWriter.set_row_range_height/4` now use the same units as their single-cell counterparts (`set_column_width/3`, `set_row_height/3`): Excel character units for column widths, points for row heights. Previously the range variants used pixels, which silently rendered very differently from the single-cell variants for the same numeric value. Callers passing pixel values will need to convert (rough rule of thumb: 1 character ≈ 7 pixels for default font; 1 point ≈ 1.333 pixels).

### fixes

- `XlsxWriter.write/5` with a URL value, link text, and `:format` option no longer drops the hyperlink. The previous implementation wrote the URL and then immediately overwrote the cell with plain formatted text, leaving the link text visible but not clickable.

### improvements

- Both NIFs (`write/1`, `write_with_properties/2`) are now scheduled on dirty CPU schedulers. Workbook serialization is CPU-bound and can take far longer than the BEAM's ~1ms NIF budget; running on dirty schedulers prevents the calling scheduler thread from stalling and avoids tail-latency spikes on unrelated processes in the same VM.
- Embedding an image via a binary no longer copies the buffer before handing it to `rust_xlsxwriter`. For multi-MB images this saves a full allocation and copy per cell.

## v0.8.3

### fixes

- `XlsxWriter.write/5` with a `:format` option now accepts the same value types as the unformatted path. Previously a `FunctionClauseError` was raised for `Date`, `DateTime`, `NaiveDateTime`, `Decimal`, atoms, and booleans because `write_with_format/5` only had clauses for `is_binary/1`, `is_number/1`, and (since v0.8.2) `nil`.
- `XlsxWriter.merge_range/7` with a `:format` option got the same treatment: it now accepts `nil`, `Date`, `DateTime`, `NaiveDateTime`, `Decimal`, and atoms in addition to the previously-supported binary, number, and boolean.
- Unknown value types passed to `write/5` or `merge_range/7` with a `:format` now raise `XlsxWriter.Error` with a clear message instead of a bare `FunctionClauseError`.
- `XlsxWriter.write/5` with `true` or `false` (no format) now produces a boolean cell. Previously the value was converted to the string `"true"` / `"false"` because the `is_atom/1` clause in `to_rust_val/1` caught booleans before any boolean-specific handling.

## v0.8.2

### fixes

- `XlsxWriter.write/5` with a format option now accepts `nil` values, writing an empty string cell with the format applied. Previously this raised `FunctionClauseError` from `write_with_format/5`, while the unformatted path (`write/4` without `:format`) already handled `nil` via `to_rust_val/1`. This makes the two paths consistent.

## v0.8.1

### improvements

- Allow `:decimal` 3.0 alongside 2.0 in the dependency requirement.
- Update `rust_xlsxwriter` from 0.94.0 to 0.95.0
  - Optional `zmij` feature for ~10% faster numeric writes (replaces deprecated `ryu`)
  - Fix double-escaping of URLs
  - Fix chart titles containing brackets being ignored
- Update `rustler` from 0.37.3 to 0.37.4
- Update cargo transitive dependency `hashbrown` (0.17.0 → 0.17.1)

## v0.8.0

### new features

- Add workbook document properties via `XlsxWriter.WorkbookProperties` - set author, title, subject, manager, company, category, keywords, comment, and status metadata. Pass via `XlsxWriter.generate/2` with the `:properties` option.
- Add worksheet tab colors via `XlsxWriter.set_tab_color/2` - color-code sheet tabs for visual organization in multi-sheet workbooks.
- Add column autofit via `XlsxWriter.autofit/1` - automatically adjust column widths to fit the longest content.
- Expand cell formatting options:
  - `:text_wrap` - wrap text within a cell
  - `{:valign, pos}` - vertical alignment (`:top`, `:center`, `:bottom`, `:justify`, `:distributed`)
  - `{:rotation, angle}` - rotate text (-90 to 90 degrees, or 270 for vertical stacked)
  - `:shrink` - shrink text to fit cell width
  - `{:indent, level}` - indent text by a given level

### improvements

- Update cargo transitive dependencies (hashbrown, indexmap, inventory, simd-adler32)

## v0.7.6

### new features

- Add format support for formula cells - `write_formula/5` now accepts a `format:` option for applying formatting (bold, colors, borders, etc.) to formula cells (contributed by @serpent213)

### improvements

- Update rustler from 0.37.2 to 0.37.3

## v0.7.5

### improvements

- Update rust_xlsxwriter from 0.92.3 to 0.94.0
  - Enhanced autofit support for formatted numbers and dates via optional `ssfmt` crate
  - New `set_autofit_max_row()` and `set_autofit_max_width()` methods
  - Fix XML escape-like strings (e.g. `_x1234`) being doubly escaped
  - Fix floating-point precision issue in image fitting
  - Updated `zip` dependency to v7.0

## v0.7.4

- Update Rust dependencies

## v0.7.3

### new features

- Add rich text formatting support via `write_rich_string/5` - apply different formatting to different parts of text within a single cell
  - Each segment can have its own formatting (bold, italic, colors, font size, etc.)
  - Supports cell-level formatting (alignment, borders, background) via the `format:` option
  - Perfect for scientific notation (E=mc²), chemical formulas (H₂O), and mixed formatting in reports

### improvements

- Update Rust dependencies (bumpalo, libz-rs-sys, proc-macro2, quote)

## v0.7.2

### improvements

- Update rustler dependency from 0.37.0 to 0.37.2

## v0.7.0

### new features

- Add native boolean data type support via `write_boolean/5` - write Excel TRUE/FALSE values with optional formatting
- Add URL/hyperlink support via `write_url/5` - create clickable hyperlinks with custom display text and formatting
- Add blank cell support via `write_blank/4` - pre-format cells without data
- Add freeze panes support via `freeze_panes/3` - lock rows/columns when scrolling to keep headers visible
- Add hide row/column support via `hide_row/2` and `hide_column/2` - hide specific rows or columns
- Add autofilter support via `set_autofilter/5` - add dropdown filter buttons to column headers
- Add merged cells support via `merge_range/7` - combine multiple cells into a single cell
- Add cell background colors via `{:bg_color, hex_color}` format option - set cell background colors with hex codes
- Add comprehensive font styling:
  - Font colors via `{:font_color, hex_color}` - set text color with hex codes
  - Font styles via `:italic`, `:strikethrough` - apply text decoration
  - Font sizes via `{:font_size, size}` - set font size in points
  - Font families via `{:font_name, name}` - use custom fonts (Arial, Courier, etc.)
  - Text position via `:superscript`, `:subscript` - create scientific notation and chemical formulas
  - Underline styles via `{:underline, style}` - single, double, accounting underlines
- Add comprehensive cell border support:
  - All-sides borders via `{:border, style}` - apply same border to all sides
  - Individual borders via `{:border_top, style}`, `{:border_bottom, style}`, `{:border_left, style}`, `{:border_right, style}` - control each side independently
  - Border colors via `{:border_color, hex_color}` and side-specific colors - customize border colors per side
  - 13 border styles: `:thin`, `:medium`, `:thick`, `:dashed`, `:dotted`, `:double`, `:hair`, `:medium_dashed`, `:dash_dot`, `:medium_dash_dot`, `:dash_dot_dot`, `:medium_dash_dot_dot`, `:slant_dash_dot`
- Add column and row range operations for efficient bulk sizing:
  - `set_column_range_width/4` - set width for multiple consecutive columns at once
  - `set_row_range_height/4` - set height for multiple consecutive rows at once
  - Simplifies setting uniform sizes across ranges (e.g., set columns A-E to 120 pixels)

## v0.6.0

### improvements

- Update rustler dependency from 0.36.2 to 0.37.0 - see [rust_xlsxwriter changes](https://github.com/rusterlium/rustler/blob/main/CHANGELOG.md#v0370---2025-11-22)
- Update rust_xlsxwriter dependency from 0.90.0 to 0.90.2 - see [rust_xlsxwriter changes](https://github.com/jmcnamara/rust_xlsxwriter/blob/main/CHANGELOG.md#version-0902---october-8-2024)
- Update various Elixir dependencies (igniter, ex_doc, file_system, rewrite)

## v0.5.0

### breaking

- Rename module `XlsxWriter.Workbook` to `XlsxWriter` - this simplifies the API by removing the nested module structure. All functions that were previously called as `XlsxWriter.Workbook.function_name()` should now be called as `XlsxWriter.function_name()`

### improvements

- Cleaner, more intuitive module structure
- Simplified imports - no need for `alias XlsxWriter.Workbook` anymore

## v0.4.0

### breaking

- Unify write and write_with_format functions - the API has been simplified with format options now passed as part of the options map

### improvements

- Add comprehensive formatting options documentation to README
- Update package description for better clarity
- Clean up and improve README documentation with advanced usage examples

## v0.3.6

- No changes, just a release fix

## v0.3.5

### improvements

- Clean up README documentation and formatting
- Update rust_xlsxwriter dependency from 0.88.0 to 0.90.0

## v0.3.0

### breaking

- XlsxWriter returns a binary string, not IO data anymore
