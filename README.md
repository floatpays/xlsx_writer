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

## Builder API (High-Level)

> **⚠️ Experimental Feature**: The Builder API is experimental and may change in future releases. While functional and tested, the API design may evolve based on user feedback. Use with caution in production code and expect potential breaking changes.

The `XlsxWriter.Builder` module provides a simplified API for generating Excel files without manually tracking cell positions. Perfect for quickly dumping data into spreadsheets.

### Basic Usage

```elixir
alias XlsxWriter.Builder

# Simple data export
Builder.create()
|> Builder.add_sheet("Summary")
|> Builder.add_rows([
  ["Name", "Age", "City"],
  ["Alice", 30, "NYC"],
  ["Bob", 25, "LA"]
])
|> Builder.write_file("report.xlsx")
```

### With Formatting

Format individual cells using tuples with options:

```elixir
Builder.create()
|> Builder.add_sheet("Q1 Report")
|> Builder.add_rows([
  # Header row with bold formatting and column widths
  [{"Quarter", format: [:bold], width: 15}, {"Revenue", format: [:bold], width: 15}],
  ["Q1", {170000, format: [{:num_format, "$#,##0"}]}],
  ["Q2", {185000, format: [{:num_format, "$#,##0"}]}]
])
|> Builder.skip_rows(1)  # Add spacing
|> Builder.add_rows([
  [{"Total", format: [:bold, :italic]}, {355000, format: [{:num_format, "$#,##0"}]}]
])
|> Builder.write_file("styled_report.xlsx")
```

### Multiple Sheets

Switch between sheets effortlessly:

```elixir
Builder.create()
|> Builder.add_sheet("Summary")
|> Builder.add_rows([
  [{"Total Sales", format: [:bold]}, 250000]
])
|> Builder.add_sheet("Details")
|> Builder.add_rows([
  [{"Product", format: [:bold]}, {"Amount", format: [:bold]}],
  ["Widget A", 50000],
  ["Widget B", 62000]
])
|> Builder.write_file("workbook.xlsx")
```

### Large Datasets

Ideal for exporting large amounts of data:

```elixir
# Generate 10,000 rows efficiently
data = Enum.map(1..10_000, fn i ->
  ["Record #{i}", i * 100, :rand.uniform(1000)]
end)

Builder.create()
|> Builder.add_sheet("Data")
|> Builder.add_rows([
  [{"ID", format: [:bold]}, {"Value", format: [:bold]}, {"Score", format: [:bold]}]
])
|> Builder.add_rows(data)
|> Builder.write_file("large.xlsx")
```

### Positioning Options

Override cursor position when needed:

```elixir
Builder.create()
|> Builder.add_sheet("Layout")
|> Builder.add_rows([["Top Left"]], start_row: 0, start_col: 0)
|> Builder.add_rows([["Offset Data"]], start_row: 5, start_col: 5)
|> Builder.write_file("positioned.xlsx")
```

### Available Format Options in Builder

When using tuple format `{value, opts}`, you can specify:

- `width: number` - Column width (Builder-specific, applies to entire column)
- `format: [...]` - List of XlsxWriter format options:
  - `:bold` - Bold text
  - `:italic` - Italic text
  - `:strikethrough` - Strikethrough text
  - `{:font_size, number}` - Font size in points
  - `{:font_color, "#RRGGBB"}` - Font color (hex)
  - `{:bg_color, "#RRGGBB"}` - Background color (hex)
  - `{:align, :left | :center | :right}` - Text alignment
  - `{:num_format, "format_string"}` - Number format
  - `{:border, style}` - Border style (`:thin`, `:medium`, `:thick`, etc.)
  - `{:border_top, style}`, `{:border_bottom, style}`, etc. - Individual borders

The `format` option uses the same format list as `XlsxWriter.write/5`.

### Builder Demo

Run the Builder examples to see all features:

```bash
mix run examples/builder_demo.exs
```

This creates 5 example files demonstrating:
- Simple reports with automatic positioning
- Multi-sheet workbooks with formatting
- Large dataset generation (1000+ rows)
- Complex formatting with colors and styles
- Positioned data with explicit coordinates

## Detailed Usage (Low-Level API)

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

### Font Styling

Apply comprehensive font styling with colors, sizes, styles, and text positioning:

```elixir
sheet =
  XlsxWriter.new_sheet("Typography")
  # Font colors
  |> XlsxWriter.write(0, 0, "Red Text", format: [{:font_color, "#FF0000"}])
  |> XlsxWriter.write(0, 1, "Blue Text", format: [{:font_color, "#0000FF"}])
  |> XlsxWriter.write(0, 2, "Green Text", format: [{:font_color, "#00FF00"}])

  # Font styles
  |> XlsxWriter.write(1, 0, "Italic", format: [:italic])
  |> XlsxWriter.write(1, 1, "Strikethrough", format: [:strikethrough])
  |> XlsxWriter.write(1, 2, "Underlined", format: [{:underline, :single}])

  # Font sizes
  |> XlsxWriter.write(2, 0, "Small", format: [{:font_size, 10}])
  |> XlsxWriter.write(2, 1, "Medium", format: [{:font_size, 14}])
  |> XlsxWriter.write(2, 2, "Large", format: [{:font_size, 18}])

  # Font families
  |> XlsxWriter.write(3, 0, "Arial", format: [{:font_name, "Arial"}])
  |> XlsxWriter.write(3, 1, "Courier", format: [{:font_name, "Courier New"}])
  |> XlsxWriter.write(3, 2, "Times", format: [{:font_name, "Times New Roman"}])

  # Combined formatting
  |> XlsxWriter.write(4, 0, "Bold Red Large",
      format: [:bold, {:font_color, "#FF0000"}, {:font_size, 16}])

  # Scientific notation and chemical formulas
  |> XlsxWriter.write(5, 0, "E=mc²", format: [:superscript])
  |> XlsxWriter.write(5, 1, "H₂O", format: [:subscript])

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("typography.xlsx", content)
```

Available underline styles: `:single`, `:double`, `:single_accounting`, `:double_accounting`

### Cell Borders

Add professional-looking borders to cells with various styles and colors:

```elixir
sheet =
  XlsxWriter.new_sheet("Invoice")
  # Headers with thick borders and background
  |> XlsxWriter.write(0, 0, "Item",
      format: [:bold, {:border, :thick}, {:bg_color, "#4472C4"}, {:align, :center}])
  |> XlsxWriter.write(0, 1, "Quantity",
      format: [:bold, {:border, :thick}, {:bg_color, "#4472C4"}, {:align, :center}])
  |> XlsxWriter.write(0, 2, "Price",
      format: [:bold, {:border, :thick}, {:bg_color, "#4472C4"}, {:align, :center}])
  |> XlsxWriter.write(0, 3, "Total",
      format: [:bold, {:border, :thick}, {:bg_color, "#4472C4"}, {:align, :center}])

  # Data rows with thin borders
  |> XlsxWriter.write(1, 0, "Widget A", format: [{:border, :thin}])
  |> XlsxWriter.write(1, 1, 10, format: [{:border, :thin}])
  |> XlsxWriter.write(1, 2, 99.99, format: [{:border, :thin}, {:num_format, "$#,##0.00"}])
  |> XlsxWriter.write_formula(1, 3, "=B2*C2")
  |> XlsxWriter.write(1, 3, nil, format: [{:border, :thin}, {:num_format, "$#,##0.00"}])

  # Total row with double bottom border
  |> XlsxWriter.write(2, 2, "Total:", format: [:bold, {:border_right, :thin}])
  |> XlsxWriter.write_formula(2, 3, "=D2")
  |> XlsxWriter.write(2, 3, nil,
      format: [:bold, {:border_bottom, :double}, {:num_format, "$#,##0.00"}])

  # Colored borders for emphasis
  |> XlsxWriter.write(4, 0, "Important Note",
      format: [{:border, :medium}, {:border_color, "#FF0000"}])

  # Multi-colored borders (different color per side)
  |> XlsxWriter.write(5, 0, "Rainbow Border",
      format: [
        {:border_top, :thin}, {:border_top_color, "#FF0000"},
        {:border_right, :thin}, {:border_right_color, "#00FF00"},
        {:border_bottom, :thin}, {:border_bottom_color, "#0000FF"},
        {:border_left, :thin}, {:border_left_color, "#FFFF00"}
      ])

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("invoice.xlsx", content)
```

Available border styles: `:thin`, `:medium`, `:thick`, `:dashed`, `:dotted`, `:double`, `:hair`, `:medium_dashed`, `:dash_dot`, `:medium_dash_dot`, `:dash_dot_dot`, `:medium_dash_dot_dot`, `:slant_dash_dot`

### Cell Background Colors

Add visual emphasis with cell background colors:

```elixir
sheet =
  XlsxWriter.new_sheet("Status Report")
  # Headers with background colors
  |> XlsxWriter.write(0, 0, "Status", format: [:bold, {:bg_color, "#4472C4"}])
  |> XlsxWriter.write(0, 1, "Item", format: [:bold, {:bg_color, "#4472C4"}])
  |> XlsxWriter.write(0, 2, "Value", format: [:bold, {:bg_color, "#4472C4"}])

  # Success (green)
  |> XlsxWriter.write(1, 0, "Complete", format: [{:bg_color, "#C6E0B4"}])
  |> XlsxWriter.write(1, 1, "Task A")
  |> XlsxWriter.write(1, 2, 100)

  # Warning (yellow)
  |> XlsxWriter.write(2, 0, "Pending", format: [{:bg_color, "#FFE699"}])
  |> XlsxWriter.write(2, 1, "Task B")
  |> XlsxWriter.write(2, 2, 75)

  # Error (red)
  |> XlsxWriter.write(3, 0, "Failed", format: [{:bg_color, "#F4B084"}])
  |> XlsxWriter.write(3, 1, "Task C")
  |> XlsxWriter.write(3, 2, 0)

  # Combined formatting
  |> XlsxWriter.write(4, 0, "Total",
      format: [:bold, {:align, :center}, {:bg_color, "#D9D9D9"}])

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("status_report.xlsx", content)
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

  # Set multiple columns at once (columns A-E to 120 pixels)
  |> XlsxWriter.set_column_range_width(0, 4, 120)

  # Set row height
  |> XlsxWriter.set_row_height(0, 40)

  # Set multiple rows at once (rows 1-10 to 25 pixels)
  |> XlsxWriter.set_row_range_height(1, 10, 25)

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
| **Font Weight** | `:bold` | `format: [:bold]` |
| **Font Style** | `:italic` | `format: [:italic]` |
| | `:strikethrough` | `format: [:strikethrough]` |
| **Font Color** | `{:font_color, hex}` | `format: [{:font_color, "#FF0000"}]` |
| **Font Size** | `{:font_size, points}` | `format: [{:font_size, 14}]` |
| **Font Family** | `{:font_name, name}` | `format: [{:font_name, "Arial"}]` |
| **Underline** | `{:underline, style}` | `format: [{:underline, :single}]` |
| **Text Position** | `:superscript` | `format: [:superscript]` |
| | `:subscript` | `format: [:subscript]` |
| **Background** | `{:bg_color, hex}` | `format: [{:bg_color, "#FFFF00"}]` |
| **Borders** | `{:border, style}` | `format: [{:border, :thin}]` |
| | `{:border_top, style}` | `format: [{:border_top, :thick}]` |
| | `{:border_bottom, style}` | `format: [{:border_bottom, :double}]` |
| | `{:border_left, style}` | `format: [{:border_left, :dashed}]` |
| | `{:border_right, style}` | `format: [{:border_right, :dotted}]` |
| **Border Colors** | `{:border_color, hex}` | `format: [{:border_color, "#000000"}]` |
| | `{:border_top_color, hex}` | `format: [{:border_top_color, "#FF0000"}]` |
| | `{:border_bottom_color, hex}` | `format: [{:border_bottom_color, "#00FF00"}]` |
| | `{:border_left_color, hex}` | `format: [{:border_left_color, "#0000FF"}]` |
| | `{:border_right_color, hex}` | `format: [{:border_right_color, "#FFFF00"}]` |
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

> **Note**: XlsxWriter implements comprehensive formatting including fonts, colors, borders, alignment, and number formats. Additional formatting features may be added in future releases based on user demand.

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
