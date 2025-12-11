# Advanced Formatting Guide

This guide covers all the formatting options available in XlsxWriter.

## Font Styling

Apply comprehensive font styling with colors, sizes, styles, and text positioning:

```elixir
sheet = XlsxWriter.new_sheet("Typography")
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

**Available underline styles:** `:single`, `:double`, `:single_accounting`, `:double_accounting`

## Rich Text Formatting

Apply different formatting to different parts of text within a single cell using rich strings:

```elixir
sheet = XlsxWriter.new_sheet("Rich Text")
  # Bold and normal text in the same cell
  |> XlsxWriter.write_rich_string(0, 0, [
    {"Bold ", [:bold]},
    {"Normal ", []},
    {"Italic", [:italic]}
  ])

  # Colored text segments
  |> XlsxWriter.write_rich_string(1, 0, [
    {"Red ", [{:font_color, "#FF0000"}]},
    {"Green ", [{:font_color, "#00FF00"}]},
    {"Blue", [{:font_color, "#0000FF"}]}
  ])

  # Multiple format options per segment
  |> XlsxWriter.write_rich_string(2, 0, [
    {"Bold Red ", [:bold, {:font_color, "#FF0000"}]},
    {"Large ", [{:font_size, 16}]},
    {"Underlined", [{:underline, :single}]}
  ])

  # Scientific notation with proper formatting
  |> XlsxWriter.write_rich_string(3, 0, [
    {"E=mc", []},
    {"2", [:superscript]}
  ])

  # Chemical formulas
  |> XlsxWriter.write_rich_string(4, 0, [
    {"H", []},
    {"2", [:subscript]},
    {"O", []}
  ])

  # With cell-level formatting (centered with background)
  |> XlsxWriter.write_rich_string(5, 0, [
    {"Important: ", [:bold]},
    {"Read carefully", [:italic]}
  ], format: [{:align, :center}, {:bg_color, "#FFFF00"}])

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("rich_text.xlsx", content)
```

### Rich String Segment Format Options

Each segment in a rich string can have these text formatting options:

| Option | Example |
|--------|---------|
| `:bold` | `{"Bold text", [:bold]}` |
| `:italic` | `{"Italic text", [:italic]}` |
| `:strikethrough` | `{"Struck text", [:strikethrough]}` |
| `:superscript` | `{"2", [:superscript]}` |
| `:subscript` | `{"2", [:subscript]}` |
| `{:font_color, hex}` | `{"Red", [{:font_color, "#FF0000"}]}` |
| `{:font_size, points}` | `{"Large", [{:font_size, 18}]}` |
| `{:font_name, name}` | `{"Arial", [{:font_name, "Arial"}]}` |
| `{:underline, style}` | `{"Underlined", [{:underline, :single}]}` |

The optional `format:` option applies cell-level formatting (alignment, borders, background) to the entire cell.

## Cell Borders

Add professional-looking borders to cells with various styles and colors:

```elixir
sheet = XlsxWriter.new_sheet("Invoice")
  # Headers with thick borders and background
  |> XlsxWriter.write(0, 0, "Item",
      format: [:bold, {:border, :thick}, {:bg_color, "#4472C4"}, {:align, :center}])
  |> XlsxWriter.write(0, 1, "Quantity",
      format: [:bold, {:border, :thick}, {:bg_color, "#4472C4"}, {:align, :center}])
  |> XlsxWriter.write(0, 2, "Price",
      format: [:bold, {:border, :thick}, {:bg_color, "#4472C4"}, {:align, :center}])

  # Data rows with thin borders
  |> XlsxWriter.write(1, 0, "Widget A", format: [{:border, :thin}])
  |> XlsxWriter.write(1, 1, 10, format: [{:border, :thin}])
  |> XlsxWriter.write(1, 2, 99.99, format: [{:border, :thin}, {:num_format, "$#,##0.00"}])

  # Total row with double bottom border
  |> XlsxWriter.write(2, 1, "Total:", format: [:bold, {:border_right, :thin}])
  |> XlsxWriter.write(2, 2, 999.90,
      format: [:bold, {:border_bottom, :double}, {:num_format, "$#,##0.00"}])

  # Colored borders
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

**Available border styles:** `:thin`, `:medium`, `:thick`, `:dashed`, `:dotted`, `:double`, `:hair`, `:medium_dashed`, `:dash_dot`, `:medium_dash_dot`, `:dash_dot_dot`, `:medium_dash_dot_dot`, `:slant_dash_dot`

## Cell Background Colors

Add visual emphasis with cell background colors:

```elixir
sheet = XlsxWriter.new_sheet("Status Report")
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

## Number Formatting

Apply custom number formats to cells:

```elixir
sheet = XlsxWriter.new_sheet("Formatted Numbers")
  # Currency format
  |> XlsxWriter.write(0, 0, 1234.56, format: [{:num_format, "[$R] #,##0.00"}])
  # Thousands separator
  |> XlsxWriter.write(1, 0, 98765, format: [{:num_format, "0,000.00"}])
  # Percentage
  |> XlsxWriter.write(2, 0, 0.75, format: [{:num_format, "0.00%"}])

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("formatted_numbers.xlsx", content)
```

### Common Number Format Strings

| Format | Description | Example Output |
|--------|-------------|----------------|
| `"#,##0.00"` | Thousands separator with 2 decimals | `1,234.56` |
| `"$#,##0.00"` | Currency (USD) | `$1,234.56` |
| `"0.00%"` | Percentage | `12.34%` |
| `"0.000E+00"` | Scientific notation | `1.235E+03` |
| `"mm/dd/yyyy"` | Date format | `12/25/2023` |
| `"h:mm AM/PM"` | Time format | `2:30 PM` |

## Combining Multiple Formats

You can combine multiple formatting options:

```elixir
sheet = XlsxWriter.new_sheet("Combined")
  |> XlsxWriter.write(0, 0, "Fancy Text",
      format: [
        :bold,
        :italic,
        {:font_color, "#FF0000"},
        {:font_size, 16},
        {:bg_color, "#FFFF00"},
        {:border, :thick},
        {:align, :center}
      ])

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("combined.xlsx", content)
```

## Format Options Reference

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
