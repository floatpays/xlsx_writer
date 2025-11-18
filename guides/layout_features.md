# Layout Features Guide

This guide covers advanced layout features for organizing and structuring your spreadsheets.

## Freeze Panes

Lock rows and/or columns when scrolling to keep headers visible:

```elixir
sheet = XlsxWriter.new_sheet("Sales Data")
  # Header row
  |> XlsxWriter.write(0, 0, "Product", format: [:bold])
  |> XlsxWriter.write(0, 1, "Q1", format: [:bold])
  |> XlsxWriter.write(0, 2, "Q2", format: [:bold])
  |> XlsxWriter.write(0, 3, "Q3", format: [:bold])
  |> XlsxWriter.write(0, 4, "Q4", format: [:bold])

  # Freeze first row (headers stay visible when scrolling down)
  |> XlsxWriter.freeze_panes(1, 0)

  # Data rows
  |> XlsxWriter.write(1, 0, "Widget A")
  |> XlsxWriter.write(1, 1, 100)
  |> XlsxWriter.write(1, 2, 150)
  # ... more data

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("frozen_headers.xlsx", content)
```

**Freeze options:**
- `freeze_panes(sheet, 1, 0)` - Freeze first row
- `freeze_panes(sheet, 0, 1)` - Freeze first column
- `freeze_panes(sheet, 1, 1)` - Freeze first row and column

## Merged Cells

Combine multiple cells into a single cell:

```elixir
sheet = XlsxWriter.new_sheet("Report")
  # Merged title spanning columns A-E
  |> XlsxWriter.merge_range(0, 0, 0, 4, "Q1 Sales Report",
      format: [:bold, {:align, :center}, {:font_size, 16}])

  # Column headers
  |> XlsxWriter.write(1, 0, "Product", format: [:bold])
  |> XlsxWriter.write(1, 1, "Units", format: [:bold])
  |> XlsxWriter.write(1, 2, "Price", format: [:bold])
  |> XlsxWriter.write(1, 3, "Total", format: [:bold])
  |> XlsxWriter.write(1, 4, "Status", format: [:bold])

  # Merge cells vertically for a multi-row value
  |> XlsxWriter.merge_range(2, 0, 4, 0, "Product Category A")

  # Merge cells for a 2D range
  |> XlsxWriter.merge_range(6, 1, 7, 2, "Special Note",
      format: [{:align, :center}, {:bg_color, "#FFFF00"}])

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("merged_cells.xlsx", content)
```

## Autofilter

Add dropdown filter buttons to column headers:

```elixir
sheet = XlsxWriter.new_sheet("Data")
  # Headers
  |> XlsxWriter.write(0, 0, "Name", format: [:bold])
  |> XlsxWriter.write(0, 1, "Age", format: [:bold])
  |> XlsxWriter.write(0, 2, "City", format: [:bold])
  |> XlsxWriter.write(0, 3, "Department", format: [:bold])

  # Add autofilter to header row (columns A-D)
  |> XlsxWriter.set_autofilter(0, 0, 0, 3)

  # Data rows
  |> XlsxWriter.write(1, 0, "Alice")
  |> XlsxWriter.write(1, 1, 30)
  |> XlsxWriter.write(1, 2, "NYC")
  |> XlsxWriter.write(1, 3, "Engineering")
  # ... more data

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("filterable_data.xlsx", content)
```

## Hide Rows and Columns

Hide specific rows or columns without deleting them:

```elixir
sheet = XlsxWriter.new_sheet("Hidden Data")
  # Visible row
  |> XlsxWriter.write(0, 0, "Visible Row")

  # Hidden row (row 1)
  |> XlsxWriter.write(1, 0, "This row is hidden")
  |> XlsxWriter.hide_row(1)

  # Visible columns
  |> XlsxWriter.write(0, 0, "Col A")
  |> XlsxWriter.write(0, 1, "Col B")
  |> XlsxWriter.write(0, 2, "Hidden Col")

  # Hidden column (column C)
  |> XlsxWriter.hide_column(2)

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("hidden_data.xlsx", content)
```

**Use cases for hiding:**
- Internal calculations not meant for viewing
- Template fields that should be hidden in final reports
- Temporary data for development

## Column Widths and Row Heights

Customize the size of columns and rows:

```elixir
sheet = XlsxWriter.new_sheet("Sized")
  # Set column widths (in characters)
  |> XlsxWriter.set_column_width(0, 30)  # Wide column for long text
  |> XlsxWriter.set_column_width(1, 15)  # Medium column
  |> XlsxWriter.set_column_width(2, 10)  # Narrow column

  # Set row heights (in points)
  |> XlsxWriter.set_row_height(0, 40)    # Tall header row
  |> XlsxWriter.set_row_height(1, 25)    # Medium row

  # Add data
  |> XlsxWriter.write(0, 0, "Long Description Text", format: [:bold])
  |> XlsxWriter.write(0, 1, "Name", format: [:bold])
  |> XlsxWriter.write(0, 2, "Code", format: [:bold])

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("custom_sizes.xlsx", content)
```

## Images

Embed images directly into spreadsheets:

```elixir
# From file
image_data = File.read!("logo.png")

sheet = XlsxWriter.new_sheet("With Image")
  |> XlsxWriter.write(0, 0, "Company Logo:")
  |> XlsxWriter.write_image(0, 1, image_data)
  |> XlsxWriter.set_column_width(1, 20)
  |> XlsxWriter.set_row_height(0, 80)

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("with_image.xlsx", content)
```

## Complete Layout Example

Here's a comprehensive example combining multiple layout features:

```elixir
sheet = XlsxWriter.new_sheet("Sales Report")
  # Merged header spanning columns A-E
  |> XlsxWriter.merge_range(0, 0, 0, 4, "Q1 Sales Report",
      format: [:bold, {:align, :center}, {:font_size, 16}, {:bg_color, "#4472C4"}])

  # Column headers with bold formatting and autofilter
  |> XlsxWriter.write(1, 0, "Product", format: [:bold, {:border, :thin}])
  |> XlsxWriter.write(1, 1, "Units", format: [:bold, {:border, :thin}])
  |> XlsxWriter.write(1, 2, "Price", format: [:bold, {:border, :thin}])
  |> XlsxWriter.write(1, 3, "Total", format: [:bold, {:border, :thin}])
  |> XlsxWriter.write(1, 4, "Status", format: [:bold, {:border, :thin}])
  |> XlsxWriter.set_autofilter(1, 0, 1, 4)

  # Freeze the first two rows (title + headers)
  |> XlsxWriter.freeze_panes(2, 0)

  # Set column widths
  |> XlsxWriter.set_column_width(0, 20)  # Product name
  |> XlsxWriter.set_column_width(1, 10)  # Units
  |> XlsxWriter.set_column_width(2, 12)  # Price
  |> XlsxWriter.set_column_width(3, 12)  # Total
  |> XlsxWriter.set_column_width(4, 15)  # Status

  # Data rows
  |> XlsxWriter.write(2, 0, "Widget A", format: [{:border, :thin}])
  |> XlsxWriter.write(2, 1, 150, format: [{:border, :thin}])
  |> XlsxWriter.write(2, 2, 9.99, format: [{:border, :thin}, {:num_format, "$#,##0.00"}])
  |> XlsxWriter.write_formula(2, 3, "=B3*C3")
  |> XlsxWriter.write(2, 3, nil, format: [{:border, :thin}, {:num_format, "$#,##0.00"}])
  |> XlsxWriter.write(2, 4, "Active", format: [{:border, :thin}])

  # Hidden row for internal calculations
  |> XlsxWriter.write(3, 0, "Internal Note")
  |> XlsxWriter.hide_row(3)

  # Hidden column for calculations
  |> XlsxWriter.write(2, 5, "Hidden Calc")
  |> XlsxWriter.hide_column(5)

{:ok, content} = XlsxWriter.generate([sheet])
File.write!("comprehensive_report.xlsx", content)
```
