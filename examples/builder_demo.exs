#!/usr/bin/env elixir

# This example demonstrates the high-level Builder API for XlsxWriter
# Run with: mix run examples/builder_demo.exs

alias XlsxWriter.Builder

# Example 1: Simple report with automatic positioning
IO.puts("Creating simple report...")

Builder.create()
|> Builder.add_sheet("Sales Report")
|> Builder.add_rows([
  [{"Month", format: [:bold], width: 15}, {"Revenue", format: [:bold], width: 15}],
  ["January", 50000],
  ["February", 62000],
  ["March", 58000]
])
|> Builder.skip_rows(1)
|> Builder.add_rows([[{"Total", format: [:bold]}, {170000, format: [{:num_format, "$#,##0"}]}]])
|> Builder.write_file("examples/output/simple_report.xlsx")

IO.puts("✓ Created simple_report.xlsx")

# Example 2: Multi-sheet workbook with formatting
IO.puts("\nCreating multi-sheet workbook...")

Builder.create()
|> Builder.add_sheet("Summary")
|> Builder.add_rows([
  [{"Quarterly Summary", format: [:bold, {:font_size, 16}], width: 20}]
])
|> Builder.skip_rows(1)
|> Builder.add_rows([
  [{"Quarter", format: [:bold]}, {"Revenue", format: [:bold]}, {"Expenses", format: [:bold]}, {"Profit", format: [:bold]}],
  ["Q1", 170000, 120000, 50000],
  ["Q2", 185000, 125000, 60000],
  ["Q3", 195000, 130000, 65000],
  ["Q4", 210000, 135000, 75000]
])
|> Builder.skip_rows(1)
|> Builder.add_rows([
  [
    {"Total", format: [:bold, {:bg_color, "#CCCCCC"}]},
    {760000, format: [:bold, {:num_format, "$#,##0"}]},
    {510000, format: [:bold, {:num_format, "$#,##0"}]},
    {250000, format: [:bold, {:num_format, "$#,##0"}]}
  ]
])
|> Builder.add_sheet("Q1 Details")
|> Builder.add_rows([
  [{"Product", format: [:bold], width: 25}, {"Units Sold", format: [:bold], width: 12}, {"Revenue", format: [:bold], width: 15}],
  ["Widget A", 1200, 36000],
  ["Widget B", 850, 42500],
  ["Widget C", 2100, 52500],
  ["Widget D", 650, 39000]
])
|> Builder.add_sheet("Q2 Details")
|> Builder.add_rows([
  [{"Product", format: [:bold], width: 25}, {"Units Sold", format: [:bold], width: 12}, {"Revenue", format: [:bold], width: 15}],
  ["Widget A", 1350, 40500],
  ["Widget B", 920, 46000],
  ["Widget C", 2250, 56250],
  ["Widget D", 700, 42000]
])
|> Builder.write_file("examples/output/quarterly_report.xlsx")

IO.puts("✓ Created quarterly_report.xlsx with multiple sheets")

# Example 3: Large dataset generation
IO.puts("\nCreating large dataset...")

# Generate 1000 rows of data
large_data =
  Enum.map(1..1000, fn i ->
    [
      "Product #{i}",
      :rand.uniform(1000),
      :rand.uniform(10000) / 100
    ]
  end)

Builder.create()
|> Builder.add_sheet("Products")
|> Builder.add_rows([
  [{"Product Name", format: [:bold], width: 20}, {"Quantity", format: [:bold], width: 12}, {"Price", format: [:bold], width: 15}]
])
|> Builder.add_rows(large_data)
|> Builder.write_file("examples/output/large_dataset.xlsx")

IO.puts("✓ Created large_dataset.xlsx with 1000 rows")

# Example 4: Complex formatting
IO.puts("\nCreating formatted report...")

Builder.create()
|> Builder.add_sheet("Styled Data")
|> Builder.add_rows([
  [{"Company Report", format: [:bold, {:font_size, 18}, {:font_color, "#0066CC"}], width: 25}]
])
|> Builder.skip_rows(2)
|> Builder.add_rows([
  [
    {"Status", format: [:bold, {:bg_color, "#EEEEEE"}]},
    {"Items", format: [:bold, {:bg_color, "#EEEEEE"}]},
    {"Value", format: [:bold, {:bg_color, "#EEEEEE"}]}
  ],
  [
    {"Completed", format: [{:font_color, "#00AA00"}]},
    42,
    {1250.50, format: [{:num_format, "$#,##0.00"}]}
  ],
  [
    {"In Progress", format: [{:font_color, "#FF9900"}]},
    28,
    {875.25, format: [{:num_format, "$#,##0.00"}]}
  ],
  [
    {"Pending", format: [{:font_color, "#0066CC"}]},
    15,
    {420.75, format: [{:num_format, "$#,##0.00"}]}
  ]
])
|> Builder.skip_rows(1)
|> Builder.add_rows([
  [
    {"Total", format: [:bold, :italic, {:bg_color, "#FFFF99"}]},
    {85, format: [:bold]},
    {2546.50, format: [:bold, {:num_format, "$#,##0.00"}]}
  ]
])
|> Builder.write_file("examples/output/formatted_report.xlsx")

IO.puts("✓ Created formatted_report.xlsx with colors and styles")

# Example 5: Positioned data
IO.puts("\nCreating report with positioned data...")

Builder.create()
|> Builder.add_sheet("Dashboard")
|> Builder.add_rows([
  [{"Top Section", format: [:bold]}]
], start_row: 0, start_col: 0)
|> Builder.add_rows([
  [{"Metric 1", format: [:bold]}, 100],
  [{"Metric 2", format: [:bold]}, 200]
], start_row: 2, start_col: 0)
|> Builder.add_rows([
  [{"Side Panel", format: [:bold, {:bg_color, "#E0E0E0"}]}]
], start_row: 0, start_col: 5)
|> Builder.add_rows([
  ["Info A", 123],
  ["Info B", 456],
  ["Info C", 789]
], start_row: 2, start_col: 5)
|> Builder.write_file("examples/output/positioned_data.xlsx")

IO.puts("✓ Created positioned_data.xlsx with data in different positions")

IO.puts("\n✅ All examples completed successfully!")
IO.puts("Check the examples/output/ directory for the generated files.")
