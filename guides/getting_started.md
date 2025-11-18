# Getting Started with XlsxWriter

This guide will help you get started with XlsxWriter and show you how to use its main features.

## Basic Usage

Create a new sheet and write data to it:

```elixir
sheet = XlsxWriter.new_sheet("My Sheet")
sheet = XlsxWriter.write(sheet, 0, 0, "Hello")
sheet = XlsxWriter.write(sheet, 0, 1, "World")
{:ok, xlsx_content} = XlsxWriter.generate([sheet])
File.write!("output.xlsx", xlsx_content)
```

## Supported Data Types

XlsxWriter automatically handles various Elixir data types:

```elixir
sheet = XlsxWriter.new_sheet("Data Types")
sheet = sheet
  |> XlsxWriter.write(0, 0, "String")
  |> XlsxWriter.write(1, 0, 42)
  |> XlsxWriter.write(2, 0, 3.14)
  |> XlsxWriter.write(3, 0, Date.utc_today())
  |> XlsxWriter.write(4, 0, DateTime.utc_now())
  |> XlsxWriter.write(5, 0, Decimal.new("99.99"))
  |> XlsxWriter.write_boolean(6, 0, true)
  |> XlsxWriter.write_url(7, 0, "https://example.com")

{:ok, xlsx_content} = XlsxWriter.generate([sheet])
File.write!("data_types.xlsx", xlsx_content)
```

## Basic Formatting

Apply formatting to cells using the `:format` option:

```elixir
sheet = XlsxWriter.new_sheet("Formatted")
sheet = sheet
  |> XlsxWriter.write(0, 0, "Bold Text", format: [:bold])
  |> XlsxWriter.write(0, 1, "Centered", format: [{:align, :center}])
  |> XlsxWriter.write(0, 2, "Yellow BG", format: [{:bg_color, "#FFFF00"}])
  |> XlsxWriter.write(0, 3, 1234.56, format: [{:num_format, "$#,##0.00"}])

{:ok, xlsx_content} = XlsxWriter.generate([sheet])
File.write!("formatted.xlsx", xlsx_content)
```

## Formulas

Write Excel formulas to cells:

```elixir
sheet = XlsxWriter.new_sheet("Formulas")
sheet = sheet
  |> XlsxWriter.write(0, 0, 10)
  |> XlsxWriter.write(0, 1, 20)
  |> XlsxWriter.write_formula(0, 2, "=A1+B1")

{:ok, xlsx_content} = XlsxWriter.generate([sheet])
File.write!("formulas.xlsx", xlsx_content)
```

## Column and Row Sizing

Customize column widths and row heights:

```elixir
sheet = XlsxWriter.new_sheet("Sized")
sheet = sheet
  |> XlsxWriter.write(0, 0, "Wide Column")
  |> XlsxWriter.set_column_width(0, 25)
  |> XlsxWriter.set_row_height(0, 40)

{:ok, xlsx_content} = XlsxWriter.generate([sheet])
File.write!("sized.xlsx", xlsx_content)
```

## Multiple Sheets

Create workbooks with multiple sheets:

```elixir
sheet1 = XlsxWriter.new_sheet("First Sheet")
  |> XlsxWriter.write(0, 0, "Sheet 1 Data")

sheet2 = XlsxWriter.new_sheet("Second Sheet")
  |> XlsxWriter.write(0, 0, "Sheet 2 Data")

{:ok, xlsx_content} = XlsxWriter.generate([sheet1, sheet2])
File.write!("multi_sheet.xlsx", xlsx_content)
```

## Complete Example

Here's a comprehensive example showing various features:

```elixir
sheet = XlsxWriter.new_sheet("Sales Report")
  |> XlsxWriter.write(0, 0, "Product", format: [:bold])
  |> XlsxWriter.write(0, 1, "Quantity", format: [:bold])
  |> XlsxWriter.write(0, 2, "Price", format: [:bold])
  |> XlsxWriter.write(0, 3, "Total", format: [:bold])
  |> XlsxWriter.write(1, 0, "Widget A")
  |> XlsxWriter.write(1, 1, 100)
  |> XlsxWriter.write(1, 2, 9.99)
  |> XlsxWriter.write_formula(1, 3, "=B2*C2")
  |> XlsxWriter.set_column_width(0, 15)
  |> XlsxWriter.set_column_width(1, 12)
  |> XlsxWriter.set_column_width(2, 12)
  |> XlsxWriter.set_column_width(3, 12)

{:ok, xlsx_content} = XlsxWriter.generate([sheet])
File.write!("sales_report.xlsx", xlsx_content)
```

## Next Steps

- Learn about [Advanced Formatting](formatting.md)
- Explore [Layout Features](layout_features.md)
- Check out the [API Reference](https://hexdocs.pm/xlsx_writer) for all available functions
