# v0.7.0

## new features

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

# v0.6.0

## improvements

- Update rustler dependency from 0.36.2 to 0.37.0 - see [rust_xlsxwriter changes](https://github.com/rusterlium/rustler/blob/main/CHANGELOG.md#v0370---2025-11-22)
- Update rust_xlsxwriter dependency from 0.90.0 to 0.90.2 - see [rust_xlsxwriter changes](https://github.com/jmcnamara/rust_xlsxwriter/blob/main/CHANGELOG.md#version-0902---october-8-2024)
- Update various Elixir dependencies (igniter, ex_doc, file_system, rewrite)

# v0.5.0

## breaking

- Rename module `XlsxWriter.Workbook` to `XlsxWriter` - this simplifies the API by removing the nested module structure. All functions that were previously called as `XlsxWriter.Workbook.function_name()` should now be called as `XlsxWriter.function_name()`

## improvements

- Cleaner, more intuitive module structure
- Simplified imports - no need for `alias XlsxWriter.Workbook` anymore

# v0.4.0

## breaking

- Unify write and write_with_format functions - the API has been simplified with format options now passed as part of the options map

## improvements

- Add comprehensive formatting options documentation to README
- Update package description for better clarity
- Clean up and improve README documentation with advanced usage examples

# v0.3.6

- No changes, just a release fix

# v0.3.5

## improvements

- Clean up README documentation and formatting
- Update rust_xlsxwriter dependency from 0.88.0 to 0.90.0

# v0.3.0

## breaking

- XlsxWriter returns a binary string, not IO data anymore
