# XlsxWriter

<!-- MDOC !-->

A high-performance Elixir library for creating Excel (.xlsx) spreadsheets. Built with the powerful [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) crate via Rustler NIF, providing excellent speed and memory efficiency.

## Features

- ⚡ **Fast**: Leverages Rust for high-performance spreadsheet generation
- 🧠 **Memory efficient**: Handles large datasets without excessive memory usage
- 📊 **Rich formatting**: Support for fonts, colors, borders, alignment, number formats, and more
- 🎨 **Cell borders**: Apply borders with 13 styles and customizable colors per side
- 🖼️ **Images**: Embed images directly into spreadsheets
- 📐 **Layout control**: Set column widths, row heights, bulk sizing for ranges, freeze panes, hide rows/columns
- 🧮 **Formulas**: Write Excel formulas and functions
- 🔗 **Hyperlinks**: Create clickable URLs with custom display text
- ✅ **Booleans**: Native Excel TRUE/FALSE values
- 🔀 **Merged cells**: Combine multiple cells into one
- 🔍 **Autofilter**: Add dropdown filters to headers
- ❄️ **Freeze panes**: Lock headers when scrolling
- 📄 **Multiple sheets**: Create workbooks with multiple worksheets
- 🔧 **Simple API**: Clean, pipe-friendly Elixir interface

## Quick Start

### Low-Level API

For precise control over cell positioning and advanced features:

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

### Simple API (Builder) - ⚠️ Experimental

For quickly generating files without manually tracking cell positions:

```elixir
alias XlsxWriter.Builder

Builder.create()
|> Builder.add_sheet("Sales Data")
|> Builder.add_rows([
  [{"Product", format: [:bold]}, {"Sales", format: [:bold]}, {"In Stock", format: [:bold]}],
  ["Widget A", {1500.50, format: [{:num_format, "$#,##0.00"}]}, true]
])
|> Builder.write_file("sales.xlsx")
```

> Note: The Builder API is experimental and may change in future releases. See the [Builder API section](#builder-api-high-level) for details.

## Comprehensive Demo

Want to see all features in action? Run the comprehensive demo script that showcases every XlsxWriter capability:

```bash
mix run examples/comprehensive_demo.exs
```

This generates `comprehensive_demo.xlsx` with 8 sheets demonstrating:
- All data types (strings, numbers, dates, booleans, formulas, URLs)
- Font formatting (colors, sizes, styles, families, super/subscript)
- Cell borders (all 13 styles, colored, per-side)
- Background colors and fill patterns
- Text alignment and number formats
- Layout features (freeze panes, autofilter, hidden rows/columns, range operations)
- Merged cells
- A complete invoice example

Perfect for learning the library or as a reference!

## Builder API (High-Level) - ⚠️ Experimental

The `XlsxWriter.Builder` module provides a simplified API for generating Excel files without manually tracking cell positions. Perfect for quickly dumping data into spreadsheets.

**Quick example:**

```elixir
alias XlsxWriter.Builder

Builder.create()
|> Builder.add_sheet("Summary")
|> Builder.add_rows([
  [{"Name", format: [:bold]}, {"Age", format: [:bold]}],
  ["Alice", 30],
  ["Bob", 25]
])
|> Builder.write_file("report.xlsx")
```

**Try it out:**

```bash
mix run examples/builder_demo.exs
```

This creates 5 example files showing formatting, multi-sheet workbooks, large datasets (1000+ rows), and more.

**Full documentation:** See the [Builder API Guide](guides/builder_api.md) for complete documentation, API reference, and examples.

> **Note**: The Builder API is experimental and may change in future releases.

## Guides

For detailed documentation and examples:

- **[Getting Started](guides/getting_started.md)** - Basic usage, data types, and your first spreadsheet
- **[Builder API](guides/builder_api.md)** - High-level API for quick data export (⚠️ Experimental)
- **[Advanced Formatting](guides/formatting.md)** - Fonts, colors, borders, and number formats
- **[Layout Features](guides/layout_features.md)** - Freeze panes, merged cells, autofilters, and more

## Examples

### Basic Example

```elixir
sheet =
  XlsxWriter.new_sheet("Report")
  |> XlsxWriter.write(0, 0, "Name", format: [:bold])
  |> XlsxWriter.write(0, 1, "Value", format: [:bold])
  |> XlsxWriter.write(1, 0, "Item A")
  |> XlsxWriter.write(1, 1, 100)

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("report.xlsx", content)
```

### Multi-Sheet Example

```elixir
summary = XlsxWriter.new_sheet("Summary")
  |> XlsxWriter.write(0, 0, "Total", format: [:bold])
  |> XlsxWriter.write(0, 1, 1000)

details = XlsxWriter.new_sheet("Details")
  |> XlsxWriter.write(0, 0, "Item", format: [:bold])
  |> XlsxWriter.write(1, 0, "Widget")

{:ok, content} = XlsxWriter.generate([summary, details])
File.write!("workbook.xlsx", content)
```

For more examples covering data types, formatting, layouts, and advanced features, see the [guides](#guides).

## Formatting Options

XlsxWriter supports extensive cell formatting through the `format:` parameter:

```elixir
# Font styles
format: [:bold]
format: [:italic, :strikethrough]
format: [{:font_color, "#FF0000"}, {:font_size, 14}]

# Alignment and background
format: [{:align, :center}]
format: [{:bg_color, "#FFFF00"}]

# Borders
format: [{:border, :thin}]
format: [{:border_top, :thick}, {:border_bottom, :double}]

# Number formats
format: [{:num_format, "$#,##0.00"}]  # Currency
format: [{:num_format, "0.00%"}]      # Percentage

# Combine multiple options
format: [:bold, {:align, :center}, {:bg_color, "#4472C4"}]
```

**Supported formatting:**
- **Fonts**: bold, italic, strikethrough, font color, font size, font family, underline, superscript, subscript
- **Alignment**: left, center, right
- **Colors**: background colors, font colors, border colors
- **Borders**: 13 styles (thin, medium, thick, dashed, dotted, double, etc.) for all sides
- **Numbers**: currency, percentage, thousands separator, decimals, dates, times, custom formats

For complete formatting documentation with examples, see the [Advanced Formatting Guide](guides/formatting.md).

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
