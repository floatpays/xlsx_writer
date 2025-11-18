# Demo: Cell Comments Feature
#
# This script demonstrates how to add comments/notes to cells in Excel spreadsheets.
# Comments appear when hovering over cells and are useful for documentation,
# instructions, or additional context about cell values.

alias XlsxWriter

# Create output directory if it doesn't exist
File.mkdir_p!("examples/output")

IO.puts("📝 Generating Excel file with cell comments...")

sheet =
  XlsxWriter.new_sheet("Comments Demo")
  # Title
  |> XlsxWriter.write(0, 0, "Cell Comments Demonstration",
    format: [:bold, {:font_size, 16}, {:font_color, "#0066CC"}]
  )
  |> XlsxWriter.write_comment(0, 0, "This title explains the purpose of this spreadsheet")

  # Headers with comments
  |> XlsxWriter.write(2, 0, "Product", format: [:bold, {:bg_color, "#E7E6E6"}])
  |> XlsxWriter.write_comment(2, 0, "Product name or SKU identifier")

  |> XlsxWriter.write(2, 1, "Price", format: [:bold, {:bg_color, "#E7E6E6"}])
  |> XlsxWriter.write_comment(2, 1, "Retail price in USD", author: "Finance Team")

  |> XlsxWriter.write(2, 2, "Stock", format: [:bold, {:bg_color, "#E7E6E6"}])
  |> XlsxWriter.write_comment(2, 2, "Current inventory count", author: "Warehouse")

  |> XlsxWriter.write(2, 3, "Status", format: [:bold, {:bg_color, "#E7E6E6"}])
  |> XlsxWriter.write_comment(2, 3, "Active = Currently selling\nDiscontinued = No longer available",
    author: "Product Manager"
  )

  # Data rows with contextual comments
  |> XlsxWriter.write(3, 0, "Widget A")
  |> XlsxWriter.write(3, 1, 29.99, format: [{:num_format, "$#,##0.00"}])
  |> XlsxWriter.write(3, 2, 150)
  |> XlsxWriter.write(3, 3, "Active")

  |> XlsxWriter.write(4, 0, "Gadget B")
  |> XlsxWriter.write(4, 1, 49.99, format: [{:num_format, "$#,##0.00"}])
  |> XlsxWriter.write(4, 2, 5)
  |> XlsxWriter.write_comment(4, 2, "LOW STOCK - Reorder needed by end of week",
    author: "Inventory Manager",
    visible: true,  # Make this comment always visible
    width: 250,
    height: 100
  )
  |> XlsxWriter.write(4, 3, "Active")

  |> XlsxWriter.write(5, 0, "Thingamajig C")
  |> XlsxWriter.write(5, 1, 99.99, format: [{:num_format, "$#,##0.00"}])
  |> XlsxWriter.write(5, 2, 0)
  |> XlsxWriter.write_comment(5, 2, "Out of stock - Expected delivery: Next Monday")
  |> XlsxWriter.write(5, 3, "Active")

  |> XlsxWriter.write(6, 0, "Doohickey D")
  |> XlsxWriter.write(6, 1, 19.99, format: [{:num_format, "$#,##0.00"}])
  |> XlsxWriter.write(6, 2, 200)
  |> XlsxWriter.write(6, 3, "Discontinued")
  |> XlsxWriter.write_comment(6, 3, "Discontinued on 2025-01-15\nClearance sale - 50% off",
    author: "Sales"
  )

  # Summary section
  |> XlsxWriter.write(8, 0, "Summary", format: [:bold, {:font_size, 12}])

  |> XlsxWriter.write(9, 0, "Total Items:")
  |> XlsxWriter.write(9, 1, 4)

  |> XlsxWriter.write(10, 0, "Items in Stock:")
  |> XlsxWriter.write(10, 1, 3)
  |> XlsxWriter.write_comment(10, 1, "Excludes discontinued and out-of-stock items")

  # Instructions section with large visible comment
  |> XlsxWriter.write(12, 0, "Instructions",
    format: [:bold, {:font_size, 12}, {:font_color, "#FF0000"}]
  )
  |> XlsxWriter.write_comment(12, 0,
    "HOW TO USE THIS SPREADSHEET:\n\n" <>
    "1. Hover over cells with red triangles to see comments\n" <>
    "2. Comments marked 'visible' are always shown\n" <>
    "3. Update stock levels daily before 5 PM\n" <>
    "4. Mark items as 'Discontinued' when EOL announced",
    author: "Operations Manager",
    visible: true,
    width: 350,
    height: 180
  )

  # Format columns
  |> XlsxWriter.set_column_width(0, 20)  # Product
  |> XlsxWriter.set_column_width(1, 12)  # Price
  |> XlsxWriter.set_column_width(2, 10)  # Stock
  |> XlsxWriter.set_column_width(3, 15)  # Status

case XlsxWriter.generate([sheet]) do
  {:ok, content} ->
    output_path = "examples/output/comments_demo.xlsx"
    File.write!(output_path, content)
    IO.puts("✅ Created: #{output_path}")
    IO.puts("\nFeatures demonstrated:")
    IO.puts("  ✓ Simple hover comments")
    IO.puts("  ✓ Comments with author attribution")
    IO.puts("  ✓ Always-visible comments")
    IO.puts("  ✓ Custom comment sizes")
    IO.puts("  ✓ Multi-line comment text")
    IO.puts("  ✓ Comments combined with cell formatting")
    IO.puts("\nOpen the file and hover over cells to see the comments!")

  {:error, reason} ->
    IO.puts("❌ Error: #{reason}")
    System.halt(1)
end
