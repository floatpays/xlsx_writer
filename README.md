# XlsxWriter

<!-- MDOC !-->

A high-performance Elixir library for creating Excel (.xlsx) spreadsheets. Built with the powerful [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) crate via Rustler NIF, providing excellent speed and memory efficiency.

## Features

- ⚡ **Fast**: Leverages Rust for high-performance spreadsheet generation
- 🧠 **Memory efficient**: Handles large datasets without excessive memory usage
- 📊 **Rich formatting**: Support for fonts, colors, alignment, number formats, and more
- 🖼️ **Images**: Embed images directly into spreadsheets
- 📐 **Layout control**: Set column widths, row heights, freeze panes, hide rows/columns
- 🧮 **Formulas**: Write Excel formulas and functions
- 🔗 **Hyperlinks**: Create clickable URLs with custom display text
- ✅ **Booleans**: Native Excel TRUE/FALSE values
- 🔀 **Merged cells**: Combine multiple cells into one
- 🔍 **Autofilter**: Add dropdown filters to headers
- ❄️ **Freeze panes**: Lock headers when scrolling
- 📄 **Multiple sheets**: Create workbooks with multiple worksheets
- 🔧 **Simple API**: Clean, pipe-friendly Elixir interface

## Quick Start

```elixir
# Create a simple spreadsheet
sheet =
  XlsxWriter.new_sheet("Sales Data")
  |> XlsxWriter.write(0, 0, "Product", format: [:bold])
  |> XlsxWriter.write(0, 1, "Sales", format: [:bold])
  |> XlsxWriter.write(0, 2, "In Stock", format: [:bold])
  |> XlsxWriter.write(1, 0, "Widget A")
  |> XlsxWriter.write(1, 1, 1500.50, format: [{:num_format, "$#,##0.00"}])
  |> XlsxWriter.write_boolean(1, 2, true)

{:ok, content} = XlsxWriter.generate([sheet])

File.write!("sales.xlsx", content)
```

## Detailed Usage

```elixir
filename = "test2.xlsx"

sheet1 =
  XlsxWriter.new_sheet("sheet number one")
  |> XlsxWriter.write(0, 0, "col1", format: [:bold])
  |> XlsxWriter.write(0, 1, "col2", format: [:bold, {:align, :center}])
  |> XlsxWriter.write(0, 2, "col3", format: [:bold, {:align, :right}])
  |> XlsxWriter.write(0, 3, nil)
  |> XlsxWriter.set_column_width(0, 40)
  |> XlsxWriter.set_column_width(3, 60)
  |> XlsxWriter.write(1, 0, "row 2 col 1")
  |> XlsxWriter.write(1, 1, 1.0)
  |> XlsxWriter.write_formula(1, 2, "=B2 + 2")
  |> XlsxWriter.write_formula(2, 1, "=PI()")
  |> XlsxWriter.write_image(3, 0, File.read!("bird.jpeg"))
  |> XlsxWriter.write(4, 3, 1)
  |> XlsxWriter.write(5, 3, DateTime.utc_now())
  |> XlsxWriter.write(6, 3, NaiveDateTime.utc_now())
  |> XlsxWriter.write(7, 3, Date.utc_today())

sheet2 =
  XlsxWriter.new_sheet("sheet number two")
  |> XlsxWriter.write(0, 0, "col1")

{:ok, content} = Workbook.generate([sheet1, sheet2])

File.write!(filename, content)
```

## Advanced Usage

### Data Types and Formatting

XlsxWriter supports various data types and formatting options:

```elixir
sheet =
  XlsxWriter.new_sheet("Data Types Example")
  # Strings with formatting
  |> XlsxWriter.write(0, 0, "Bold Text", format: [:bold])
  |> XlsxWriter.write(0, 1, "Centered", format: [{:align, :center}])
  |> XlsxWriter.write(0, 2, "Right Aligned", format: [{:align, :right}])

  # Numbers and decimals
  |> XlsxWriter.write(1, 0, 42)
  |> XlsxWriter.write(1, 1, 3.14159)
  |> XlsxWriter.write(1, 2, Decimal.new("99.99"))

  # Date and time
  |> XlsxWriter.write(2, 0, Date.utc_today())
  |> XlsxWriter.write(2, 1, DateTime.utc_now())
  |> XlsxWriter.write(2, 2, NaiveDateTime.utc_now())

  # Booleans
  |> XlsxWriter.write_boolean(3, 0, true)
  |> XlsxWriter.write_boolean(3, 1, false)
  |> XlsxWriter.write_boolean(3, 2, true, format: [:bold])

  # URLs and hyperlinks
  |> XlsxWriter.write_url(4, 0, "https://elixir-lang.org")
  |> XlsxWriter.write_url(4, 1, "https://hexdocs.pm", text: "Hex Docs")
  |> XlsxWriter.write_url(4, 2, "https://github.com", format: [:bold])

  # Formulas
  |> XlsxWriter.write_formula(5, 0, "=B2 * 2")
  |> XlsxWriter.write_formula(5, 1, "=PI()")
  |> XlsxWriter.write_formula(5, 2, "=TODAY()")

  # Blank cells with formatting (useful for templates)
  |> XlsxWriter.write_blank(6, 0, format: [{:align, :center}])

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("data_types.xlsx", content)
```

### Layout Features

Control worksheet layout with freeze panes, merged cells, autofilters, and hidden rows/columns:

```elixir
sheet =
  XlsxWriter.new_sheet("Sales Report")
  # Merged header spanning columns A-E
  |> XlsxWriter.merge_range(0, 0, 0, 4, "Q1 Sales Report",
      format: [:bold, {:align, :center}])

  # Column headers with bold formatting
  |> XlsxWriter.write(1, 0, "Product", format: [:bold])
  |> XlsxWriter.write(1, 1, "Units", format: [:bold])
  |> XlsxWriter.write(1, 2, "Price", format: [:bold])
  |> XlsxWriter.write(1, 3, "Total", format: [:bold])
  |> XlsxWriter.write(1, 4, "Status", format: [:bold])

  # Add dropdown filters to header row
  |> XlsxWriter.set_autofilter(1, 0, 1, 4)

  # Freeze the first two rows (title + headers)
  |> XlsxWriter.freeze_panes(2, 0)

  # Data rows
  |> XlsxWriter.write(2, 0, "Widget A")
  |> XlsxWriter.write(2, 1, 150)
  |> XlsxWriter.write(2, 2, 9.99)
  |> XlsxWriter.write_formula(2, 3, "=B3*C3")
  |> XlsxWriter.write(2, 4, "Active")

  # Hidden row for internal use
  |> XlsxWriter.write(3, 0, "Internal Note")
  |> XlsxWriter.hide_row(3)

  # Hidden column for calculations
  |> XlsxWriter.write(2, 5, "Hidden Calc")
  |> XlsxWriter.hide_column(5)

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("sales_report.xlsx", content)
```

### Booleans and URLs

Write native Excel boolean values and clickable hyperlinks:

```elixir
sheet =
  XlsxWriter.new_sheet("Links and Booleans")
  # Boolean values
  |> XlsxWriter.write(0, 0, "Active")
  |> XlsxWriter.write_boolean(0, 1, true)
  |> XlsxWriter.write(1, 0, "Disabled")
  |> XlsxWriter.write_boolean(1, 1, false, format: [:bold, {:align, :center}])

  # Hyperlinks
  |> XlsxWriter.write(2, 0, "Website")
  |> XlsxWriter.write_url(2, 1, "https://example.com")

  # URL with custom display text
  |> XlsxWriter.write(3, 0, "Documentation")
  |> XlsxWriter.write_url(3, 1, "https://hexdocs.pm/xlsx_writer", text: "View Docs")

  # URL with formatting
  |> XlsxWriter.write(4, 0, "GitHub")
  |> XlsxWriter.write_url(4, 1, "https://github.com/floatpays/xlsx_writer",
      text: "Source Code",
      format: [:bold, {:align, :center}])

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("links_and_booleans.xlsx", content)
```

### Number Formatting

Apply custom number formats to cells:

```elixir
sheet =
  XlsxWriter.new_sheet("Formatted Numbers")
  # Currency format
  |> XlsxWriter.write(0, 0, 1234.56, format: [{:num_format, "[$R] #,##0.00"}])
  # Thousands separator
  |> XlsxWriter.write(1, 0, 98765, format: [{:num_format, "0,000.00"}])
  # Percentage
  |> XlsxWriter.write(2, 0, 0.75, format: [{:num_format, "0.00%"}])

{:ok, content} = Workbook.generate([sheet])
File.write!("formatted_numbers.xlsx", content)
```

### Images and Layout

Add images and control column/row dimensions:

```elixir
image_data = File.read!("logo.png")

sheet =
  XlsxWriter.new_sheet("Layout Example")
  # Set column widths
  |> XlsxWriter.set_column_width(0, 30)
  |> XlsxWriter.set_column_width(1, 50)
  
  # Set row height
  |> XlsxWriter.set_row_height(0, 40)
  
  # Add images
  |> XlsxWriter.write_image(0, 0, image_data)
  |> XlsxWriter.write(0, 1, "Logo Description")

{:ok, content} = Workbook.generate([sheet])
File.write!("with_images.xlsx", content)
```

### Multiple Sheets

Create workbooks with multiple sheets:

```elixir
summary_sheet =
  XlsxWriter.new_sheet("Summary")
  |> XlsxWriter.write(0, 0, "Report Summary", format: [:bold])
  |> XlsxWriter.write(1, 0, "Total Records: 1000")

details_sheet =
  XlsxWriter.new_sheet("Details")
  |> XlsxWriter.write(0, 0, "ID", format: [:bold])
  |> XlsxWriter.write(0, 1, "Name", format: [:bold])
  |> XlsxWriter.write(0, 2, "Amount", format: [:bold])
  |> XlsxWriter.write(1, 0, 1)
  |> XlsxWriter.write(1, 1, "Item A")
  |> XlsxWriter.write(1, 2, 99.99)

{:ok, content} = Workbook.generate([summary_sheet, details_sheet])
File.write!("multi_sheet.xlsx", content)
```

## Formatting Options

XlsxWriter supports extensive cell formatting through the `format:` parameter. Currently implemented formatting options include:

### Font Formatting
```elixir
# Bold text
|> XlsxWriter.write(0, 0, "Bold Text", format: [:bold])
```

### Alignment
```elixir
# Text alignment options
|> XlsxWriter.write(0, 0, "Left", format: [{:align, :left}])
|> XlsxWriter.write(0, 1, "Center", format: [{:align, :center}])  
|> XlsxWriter.write(0, 2, "Right", format: [{:align, :right}])
```

### Number Formatting
```elixir
# Currency
|> XlsxWriter.write(0, 0, 1234.56, format: [{:num_format, "$#,##0.00"}])
|> XlsxWriter.write(0, 1, 1234.56, format: [{:num_format, "[$€] #,##0.00"}])

# Percentages
|> XlsxWriter.write(1, 0, 0.75, format: [{:num_format, "0.00%"}])

# Thousands separator
|> XlsxWriter.write(2, 0, 98765, format: [{:num_format, "#,##0"}])

# Custom formats
|> XlsxWriter.write(3, 0, 42, format: [{:num_format, "000.00"}])
```

### Combining Formats
```elixir
# Multiple formatting options can be combined
|> XlsxWriter.write(0, 0, "Bold & Centered", format: [:bold, {:align, :center}])
|> XlsxWriter.write(1, 0, 1500.00, format: [:bold, {:num_format, "$#,##0.00"}])
```

### Supported Format Options

| Format Type | Option | Example |
|-------------|--------|---------|
| **Font** | `:bold` | `format: [:bold]` |
| **Alignment** | `{:align, :left}` | `format: [{:align, :left}]` |
| | `{:align, :center}` | `format: [{:align, :center}]` |
| | `{:align, :right}` | `format: [{:align, :right}]` |
| **Numbers** | `{:num_format, "format_string"}` | `format: [{:num_format, "$#,##0.00"}]` |

### Common Number Format Strings

| Format | Description | Example Output |
|--------|-------------|----------------|
| `"#,##0.00"` | Thousands separator with 2 decimals | `1,234.56` |
| `"$#,##0.00"` | Currency (USD) | `$1,234.56` |
| `"0.00%"` | Percentage | `12.34%` |
| `"0.000E+00"` | Scientific notation | `1.235E+03` |
| `"mm/dd/yyyy"` | Date format | `12/25/2023` |
| `"h:mm AM/PM"` | Time format | `2:30 PM` |

> **Note**: XlsxWriter currently implements a subset of the formatting options available in the underlying `rust_xlsxwriter` library. Additional formatting features like colors, borders, and advanced font properties may be added in future releases.

## Installation

The package is available on [Hex](https://hex.pm/packages/xlsx_writer). Add `xlsx_writer` to your list of dependencies in `mix.exs`:

```elixir
def deps do
  [
    {:xlsx_writer, "~> 0.5.0"}
  ]
end
```

Then run:

```bash
mix deps.get
```

## Documentation

Full documentation is available at [HexDocs](https://hexdocs.pm/xlsx_writer).

## Development

### Publishing a new version

Follow the [rustler_precompiled guide](https://hexdocs.pm/rustler_precompiled/precompilation_guide.html):

1. Update version number in `mix.exs` and this README
2. Create and push a new tag: `git tag v0.1.x && git push origin main --tags`
3. Wait for GitHub Actions to build all NIFs
4. Download precompiled assets: `mix rustler_precompiled.download XlsxWriter.RustXlsxWriter --all`
5. Publish to Hex: `mix hex.publish`

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### Running Tests

```bash
mix test
```

### Building Documentation

```bash
mix docs
```

## Copyright and License

Copyright (c) 2025 Floatpays

This work is free. You can redistribute it and/or modify it under the
terms of the MIT License. See the [LICENSE.md](./LICENSE.md) file for more details.
