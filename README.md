# XlsxWriter

<!-- MDOC !-->

A fast Elixir library for writing Excel files using Rust.

## Usage

```elixir
filename = "test2.xlsx"

sheet1 =
  Workbook.new_sheet("sheet number one")
  |> Workbook.write(0, 0, "col1", format: [:bold])
  |> Workbook.write(0, 1, "col2", format: [:bold, {:align, :center}])
  |> Workbook.write(0, 2, "col3", format: [:bold, {:align, :right}])
  |> Workbook.write(0, 3, nil)
  |> Workbook.set_column_width(0, 40)
  |> Workbook.set_column_width(3, 60)
  |> Workbook.write(1, 0, "row 2 col 1")
  |> Workbook.write(1, 1, 1.0)
  |> Workbook.write_formula(1, 2, "=B2 + 2")
  |> Workbook.write_formula(2, 1, "=PI()")
  |> Workbook.write_image(3, 0, File.read!("bird.jpeg"))
  |> Workbook.write(4, 3, 1)
  |> Workbook.write(5, 3, DateTime.utc_now())
  |> Workbook.write(6, 3, NaiveDateTime.utc_now())
  |> Workbook.write(7, 3, Date.utc_today())

sheet2 =
  Workbook.new_sheet("sheet number two")
  |> Workbook.write(0, 0, "col1")

{:ok, content} = Workbook.generate([sheet1, sheet2])

File.write!(filename, content)
```

## Advanced Usage

### Data Types and Formatting

XlsxWriter supports various data types and formatting options:

```elixir
alias XlsxWriter.Workbook

sheet =
  Workbook.new_sheet("Data Types Example")
  # Strings with formatting
  |> Workbook.write(0, 0, "Bold Text", format: [:bold])
  |> Workbook.write(0, 1, "Centered", format: [{:align, :center}])
  |> Workbook.write(0, 2, "Right Aligned", format: [{:align, :right}])
  
  # Numbers and decimals
  |> Workbook.write(1, 0, 42)
  |> Workbook.write(1, 1, 3.14159)
  |> Workbook.write(1, 2, Decimal.new("99.99"))
  
  # Date and time
  |> Workbook.write(2, 0, Date.utc_today())
  |> Workbook.write(2, 1, DateTime.utc_now())
  |> Workbook.write(2, 2, NaiveDateTime.utc_now())
  
  # Formulas
  |> Workbook.write_formula(3, 0, "=B2 * 2")
  |> Workbook.write_formula(3, 1, "=PI()")
  |> Workbook.write_formula(3, 2, "=TODAY()")

{:ok, content} = Workbook.generate([sheet])
File.write!("data_types.xlsx", content)
```

### Number Formatting

Apply custom number formats to cells:

```elixir
sheet =
  Workbook.new_sheet("Formatted Numbers")
  # Currency format
  |> Workbook.write(0, 0, 1234.56, format: [{:num_format, "[$R] #,##0.00"}])
  # Thousands separator
  |> Workbook.write(1, 0, 98765, format: [{:num_format, "0,000.00"}])
  # Percentage
  |> Workbook.write(2, 0, 0.75, format: [{:num_format, "0.00%"}])

{:ok, content} = Workbook.generate([sheet])
File.write!("formatted_numbers.xlsx", content)
```

### Images and Layout

Add images and control column/row dimensions:

```elixir
image_data = File.read!("logo.png")

sheet =
  Workbook.new_sheet("Layout Example")
  # Set column widths
  |> Workbook.set_column_width(0, 30)
  |> Workbook.set_column_width(1, 50)
  
  # Set row height
  |> Workbook.set_row_height(0, 40)
  
  # Add images
  |> Workbook.write_image(0, 0, image_data)
  |> Workbook.write(0, 1, "Logo Description")

{:ok, content} = Workbook.generate([sheet])
File.write!("with_images.xlsx", content)
```

### Multiple Sheets

Create workbooks with multiple sheets:

```elixir
summary_sheet =
  Workbook.new_sheet("Summary")
  |> Workbook.write(0, 0, "Report Summary", format: [:bold])
  |> Workbook.write(1, 0, "Total Records: 1000")

details_sheet =
  Workbook.new_sheet("Details")
  |> Workbook.write(0, 0, "ID", format: [:bold])
  |> Workbook.write(0, 1, "Name", format: [:bold])
  |> Workbook.write(0, 2, "Amount", format: [:bold])
  |> Workbook.write(1, 0, 1)
  |> Workbook.write(1, 1, "Item A")
  |> Workbook.write(1, 2, 99.99)

{:ok, content} = Workbook.generate([summary_sheet, details_sheet])
File.write!("multi_sheet.xlsx", content)
```

## Installation

Add `xlsx_writer` to your list of dependencies in `mix.exs`:

```elixir
def deps do
  [
    {:xlsx_writer, "~> 0.3.6"}
  ]
end
```

Documentation can be found at [HexDocs](https://hexdocs.pm/xlsx_writer).

## Development

### Publishing a new version

Follow the [rustler_precompiled guide](https://hexdocs.pm/rustler_precompiled/precompilation_guide.html):

1. Update version number in `mix.exs` and this README
2. Create and push a new tag: `git tag v0.1.x && git push origin main --tags`
3. Wait for GitHub Actions to build all NIFs
4. Download precompiled assets: `mix rustler_precompiled.download XlsxWriter.RustXlsxWriter --all`
5. Publish to Hex: `mix hex.publish`

## Copyright and License

Copyright (c) 2025 Floatpays

This work is free. You can redistribute it and/or modify it under the
terms of the MIT License. See the [LICENSE.md](./LICENSE.md) file for more details.
