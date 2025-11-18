# Builder API Guide

> **⚠️ Experimental Feature**: The Builder API is experimental and may change in future releases. While functional and tested, the API design may evolve based on user feedback. Use with caution in production code and expect potential breaking changes.

The Builder API provides a high-level, simplified interface for creating Excel files without manually tracking cell positions. It's perfect for quickly generating reports and exporting data.

## When to Use the Builder API

**Use Builder when:**
- Quickly exporting data from lists or database queries
- Generating simple reports with sequential data
- You don't need advanced features like formulas, images, or merged cells
- Speed of development is more important than fine-grained control
- Working with large datasets in a straightforward layout

**Use Low-Level API when:**
- Need precise control over cell positioning
- Using advanced features (formulas, images, merged cells, freeze panes)
- Creating complex layouts with scattered data
- Building spreadsheet templates with specific structure

## Quick Start

The simplest possible example:

```elixir
alias XlsxWriter.Builder

Builder.create()
|> Builder.add_sheet("Report")
|> Builder.add_rows([
  ["Name", "Value"],
  ["Alice", 100],
  ["Bob", 200]
])
|> Builder.write_file("report.xlsx")
```

## Core Concepts

### 1. Builder State

The Builder maintains state including:
- List of sheets with their instructions
- Current active sheet
- Cursor position (row, col)
- Column width mappings

All operations return the updated builder state for piping.

### 2. Automatic Cursor Tracking

The cursor automatically advances as you add data:

```elixir
Builder.create()
|> Builder.add_sheet("Data")
|> Builder.add_rows([["Row 1"]])  # Cursor at (0,0) → (1,0)
|> Builder.add_rows([["Row 2"]])  # Cursor at (1,0) → (2,0)
|> Builder.add_rows([["Row 3"]])  # Cursor at (2,0) → (3,0)
```

### 3. Format Options

Cell formatting uses the same syntax as `XlsxWriter.write/5`:

```elixir
# Cell with formatting
{"Bold Text", format: [:bold]}

# Multiple format options
{"Styled", format: [:bold, :italic, {:font_size, 14}, {:bg_color, "#FFFF00"}]}

# Column width (Builder-specific)
{"Wide", format: [:bold], width: 30}
```

## API Reference

### `create/0`

Creates a new builder instance.

```elixir
builder = Builder.create()
```

**Returns:** A new builder state

---

### `add_sheet/2`

Adds a new sheet and switches context to it. Resets cursor to (0, 0).

```elixir
builder
|> Builder.add_sheet("Summary")
|> Builder.add_sheet("Details")
```

**Parameters:**
- `builder` - The builder state
- `sheet_name` - Name of the sheet (string)

**Returns:** Updated builder with new active sheet

---

### `add_rows/3`

Adds multiple rows starting at the current cursor position.

```elixir
# Simple rows
builder |> Builder.add_rows([
  ["A", "B", "C"],
  [1, 2, 3]
])

# With formatting
builder |> Builder.add_rows([
  [{"Header", format: [:bold]}, {"Value", format: [:bold]}],
  ["Data", 100]
])

# Override position
builder |> Builder.add_rows(
  [["Data"]],
  start_row: 5,
  start_col: 2
)
```

**Parameters:**
- `builder` - The builder state
- `rows` - List of rows (list of lists)
- `opts` - Options (optional):
  - `:start_row` - Override cursor row (0-based)
  - `:start_col` - Override cursor column (0-based)

**Returns:** Updated builder with cursor moved after last row

---

### `skip_rows/2`

Moves cursor down by N rows for spacing.

```elixir
builder
|> Builder.add_rows([["Section 1"]])
|> Builder.skip_rows(2)  # Add 2 blank rows
|> Builder.add_rows([["Section 2"]])
```

**Parameters:**
- `builder` - The builder state
- `n` - Number of rows to skip (default: 1)

**Returns:** Updated builder with cursor moved down

---

### `write_binary/1`

Generates the XLSX file as binary data.

```elixir
{:ok, content} = Builder.write_binary(builder)
File.write!("output.xlsx", content)
```

**Returns:** `{:ok, binary}` or `{:error, reason}`

---

### `write_file/2`

Generates and writes the XLSX file to disk.

```elixir
builder |> Builder.write_file("report.xlsx")
```

**Parameters:**
- `builder` - The builder state
- `path` - Output file path (required)

**Returns:** `:ok` or `{:error, reason}`

---

## Format Options

All format options from `XlsxWriter.write/5` are supported:

### Text Styles
```elixir
format: [:bold]
format: [:italic]
format: [:strikethrough]
format: [:bold, :italic]  # Combine multiple
```

### Fonts
```elixir
format: [{:font_size, 14}]
format: [{:font_color, "#FF0000"}]  # Red
format: [{:font_name, "Arial"}]
```

### Colors and Alignment
```elixir
format: [{:bg_color, "#FFFF00"}]  # Yellow background
format: [{:align, :left}]
format: [{:align, :center}]
format: [{:align, :right}]
```

### Number Formatting
```elixir
format: [{:num_format, "$#,##0.00"}]  # Currency
format: [{:num_format, "0.00%"}]       # Percentage
format: [{:num_format, "#,##0"}]       # Thousands
```

### Borders
```elixir
format: [{:border, :thin}]
format: [{:border_top, :thick}]
format: [{:border_bottom, :double}]
format: [{:border_left, :dashed}]
format: [{:border_right, :dotted}]
```

Available border styles: `:thin`, `:medium`, `:thick`, `:dashed`, `:dotted`, `:double`, `:hair`, `:medium_dashed`, `:dash_dot`, `:medium_dash_dot`, `:dash_dot_dot`, `:medium_dash_dot_dot`, `:slant_dash_dot`

### Column Width (Builder-Specific)
```elixir
{"Text", width: 20}  # Set column width
{"Text", format: [:bold], width: 30}  # With formatting
```

## Examples

### Basic Report

```elixir
alias XlsxWriter.Builder

Builder.create()
|> Builder.add_sheet("Sales Report")
|> Builder.add_rows([
  [{"Month", format: [:bold]}, {"Revenue", format: [:bold]}],
  ["January", {50000, format: [{:num_format, "$#,##0"}]}],
  ["February", {62000, format: [{:num_format, "$#,##0"}]}],
  ["March", {58000, format: [{:num_format, "$#,##0"}]}]
])
|> Builder.skip_rows(1)
|> Builder.add_rows([
  [{"Total", format: [:bold]}, {170000, format: [{:num_format, "$#,##0"}]}]
])
|> Builder.write_file("sales_report.xlsx")
```

### Multi-Sheet Workbook

```elixir
Builder.create()
|> Builder.add_sheet("Summary")
|> Builder.add_rows([
  [{"Quarterly Revenue", format: [:bold, {:font_size, 16}]}]
])
|> Builder.skip_rows(1)
|> Builder.add_rows([
  [{"Quarter", format: [:bold]}, {"Amount", format: [:bold]}],
  ["Q1", 170000],
  ["Q2", 185000],
  ["Q3", 195000],
  ["Q4", 210000]
])
|> Builder.add_sheet("Q1 Details")
|> Builder.add_rows([
  [
    {"Product", format: [:bold], width: 25},
    {"Units", format: [:bold], width: 12},
    {"Revenue", format: [:bold], width: 15}
  ],
  ["Widget A", 1200, 36000],
  ["Widget B", 850, 42500]
])
|> Builder.write_file("quarterly_report.xlsx")
```

### Large Dataset Export

```elixir
# Generate 10,000 rows from database query
data = MyApp.Repo.all(User)
|> Enum.map(fn user ->
  [user.name, user.email, user.created_at, user.status]
end)

Builder.create()
|> Builder.add_sheet("Users")
|> Builder.add_rows([
  [
    {"Name", format: [:bold], width: 20},
    {"Email", format: [:bold], width: 30},
    {"Created", format: [:bold], width: 15},
    {"Status", format: [:bold], width: 10}
  ]
])
|> Builder.add_rows(data)
|> Builder.write_file("users_export.xlsx")
```

### Styled Report

```elixir
Builder.create()
|> Builder.add_sheet("Status Report")
|> Builder.add_rows([
  [{"Project Status Dashboard", format: [:bold, {:font_size, 18}, {:font_color, "#0066CC"}]}]
])
|> Builder.skip_rows(2)
|> Builder.add_rows([
  [
    {"Status", format: [:bold, {:bg_color, "#EEEEEE"}]},
    {"Count", format: [:bold, {:bg_color, "#EEEEEE"}]},
    {"Percentage", format: [:bold, {:bg_color, "#EEEEEE"}]}
  ],
  [
    {"Completed", format: [{:font_color, "#00AA00"}]},
    42,
    {0.52, format: [{:num_format, "0.0%"}]}
  ],
  [
    {"In Progress", format: [{:font_color, "#FF9900"}]},
    28,
    {0.35, format: [{:num_format, "0.0%"}]}
  ],
  [
    {"Pending", format: [{:font_color, "#0066CC"}]},
    10,
    {0.13, format: [{:num_format, "0.0%"}]}
  ]
])
|> Builder.skip_rows(1)
|> Builder.add_rows([
  [
    {"Total", format: [:bold, {:bg_color, "#FFFF99"}]},
    {80, format: [:bold]},
    {1.0, format: [:bold, {:num_format, "0.0%"}]}
  ]
])
|> Builder.write_file("status_report.xlsx")
```

### Positioned Data

```elixir
# Create dashboard-style layout with data in specific positions
Builder.create()
|> Builder.add_sheet("Dashboard")
# Top-left section
|> Builder.add_rows([
  [{"Sales Metrics", format: [:bold, {:font_size, 14}]}]
], start_row: 0, start_col: 0)
|> Builder.add_rows([
  [{"Metric", format: [:bold]}, {"Value", format: [:bold]}],
  ["Daily Sales", 15000],
  ["Monthly Sales", 450000]
], start_row: 2, start_col: 0)
# Top-right section
|> Builder.add_rows([
  [{"Team Performance", format: [:bold, {:font_size, 14}]}]
], start_row: 0, start_col: 5)
|> Builder.add_rows([
  [{"Team", format: [:bold]}, {"Target", format: [:bold]}],
  ["East", {125, format: [{:num_format, "0%"}]}],
  ["West", {98, format: [{:num_format, "0%"}]}]
], start_row: 2, start_col: 5)
|> Builder.write_file("dashboard.xlsx")
```

## Performance Tips

1. **Batch your rows**: Add multiple rows in one `add_rows/3` call rather than multiple single-row calls
2. **Use simple values**: Plain values are faster than formatted tuples
3. **Minimize format changes**: Group cells with similar formatting together
4. **Large datasets**: The Builder handles 10,000+ rows efficiently

## Common Patterns

### Headers with Data

```elixir
# Pattern: Bold headers + regular data
Builder.add_rows([
  [{"Col1", format: [:bold]}, {"Col2", format: [:bold]}],
  ["Data1", "Data2"],
  ["Data3", "Data4"]
])
```

### Sections with Spacing

```elixir
# Pattern: Section header + spacing + data + spacing
builder
|> Builder.add_rows([[{"Section 1", format: [:bold]}]])
|> Builder.skip_rows(1)
|> Builder.add_rows(section1_data)
|> Builder.skip_rows(2)
|> Builder.add_rows([[{"Section 2", format: [:bold]}]])
|> Builder.skip_rows(1)
|> Builder.add_rows(section2_data)
```

### Totals Row

```elixir
# Pattern: Data + blank row + totals
builder
|> Builder.add_rows(data_rows)
|> Builder.skip_rows(1)
|> Builder.add_rows([
  [{"Total", format: [:bold]}, {sum, format: [:bold, {:num_format, "#,##0"}]}]
])
```

## Error Handling

```elixir
# write_binary/1 returns result tuple
case Builder.write_binary(builder) do
  {:ok, content} ->
    File.write!("output.xlsx", content)
    :ok

  {:error, reason} ->
    Logger.error("Failed to generate XLSX: #{reason}")
    {:error, reason}
end

# write_file/2 also returns result
case Builder.write_file(builder, path) do
  :ok ->
    IO.puts("✓ File created: #{path}")

  {:error, reason} ->
    IO.puts("✗ Error: #{reason}")
end
```

Common errors:
- `"No sheets added"` - Call `add_sheet/2` before adding rows
- `"No active sheet"` - Add at least one sheet with `add_sheet/2`

## Demo Scripts

Run the included demo to see all features in action:

```bash
mix run examples/builder_demo.exs
```

This generates 5 example files demonstrating:
- Simple reports with automatic positioning
- Multi-sheet workbooks with formatting
- Large dataset generation (1000 rows)
- Complex formatting with colors and styles
- Positioned data with explicit coordinates

## Comparison: Builder vs Low-Level

### Builder API
```elixir
Builder.create()
|> Builder.add_sheet("Data")
|> Builder.add_rows([
  [{"Name", format: [:bold], width: 20}, {"Age", format: [:bold]}],
  ["Alice", 30],
  ["Bob", 25]
])
|> Builder.write_file("output.xlsx")
```

### Low-Level API
```elixir
sheet = XlsxWriter.new_sheet("Data")
  |> XlsxWriter.set_column_width(0, 20)
  |> XlsxWriter.write(0, 0, "Name", format: [:bold])
  |> XlsxWriter.write(0, 1, "Age", format: [:bold])
  |> XlsxWriter.write(1, 0, "Alice")
  |> XlsxWriter.write(1, 1, 30)
  |> XlsxWriter.write(2, 0, "Bob")
  |> XlsxWriter.write(2, 1, 25)

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("output.xlsx", content)
```

The Builder API is more concise for sequential data!

## Limitations

The Builder API currently does not support:
- Formulas (use low-level API)
- Images (use low-level API)
- Merged cells (use low-level API)
- Freeze panes (use low-level API)
- Autofilter (use low-level API)
- Hide rows/columns (use low-level API)
- Row height control (use low-level API)

For these features, use the low-level `XlsxWriter` API or consider mixing both:

```elixir
# Use Builder for bulk data
builder = Builder.create()
|> Builder.add_sheet("Data")
|> Builder.add_rows(lots_of_data)

# Get the generated sheets
{:ok, binary} = Builder.write_binary(builder)

# Or build manually with low-level API for advanced features
sheet = XlsxWriter.new_sheet("Advanced")
  |> XlsxWriter.write_formula(0, 0, "=SUM(A1:A10)")
  |> XlsxWriter.freeze_panes(1, 0)

{:ok, content} = XlsxWriter.generate([sheet])
```

## Future Enhancements

Potential additions being considered:
- `skip_cols/2` - Horizontal cursor movement
- `move_to/2` - Absolute cursor positioning
- Sheet-level options (freeze panes, autofilter)
- `add_row/2` - Single row variant
- Template-based generation
- Conditional formatting helpers

Feedback welcome! This API is experimental and we want your input.
