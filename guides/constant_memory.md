# Constant memory exports

`XlsxWriter.Stream` writes rows in batches to a file. Use it when the input is
an Enumerable, such as a database stream, and the complete workbook would be
too large to hold in memory.

```elixir
rows = Stream.map(1..100_000, fn n -> [n, "Row #{n}"] end)

:ok = XlsxWriter.Stream.write_file("large.xlsx", [{"Data", rows}], batch_size: 500)
```

Each row is a list of values. A cell may be `{value, format: formats}`; the
format list has the same syntax as `XlsxWriter.write/5`. Pass multiple
`{sheet_name, rows}` pairs to create multiple worksheets. The rows Enumerable
is consumed once. Keep its source lazy to get bounded input memory use.
Use `{:formula, "=A1+B1"}` or `{:formula, "=A1+B1", format: [:bold]}` for
formula cells.

For more control, use the lower-level interface:

```elixir
{:ok, writer} = XlsxWriter.Stream.open("report.xlsx")
{:ok, writer} = XlsxWriter.Stream.add_sheet(writer, "Data")

setup =
  XlsxWriter.new_sheet("Data")
  |> XlsxWriter.set_column_width(0, 24)
  |> XlsxWriter.freeze_panes(1, 0)

{:ok, writer} = XlsxWriter.Stream.configure_sheet(writer, setup)
{:ok, writer} = XlsxWriter.Stream.write_rows(writer, [["Name"], ["Alice"]])
:ok = XlsxWriter.Stream.finish(writer)
```

`write_rows/3` accepts a bounded list. Use `start_row:` to skip forward and
`start_col:` to offset the columns. Call `abort/1` if you stop before finishing.
`write_file/3` aborts automatically when its source raises or writing fails.
The writer holds mutable native state. Use the writer returned by each call in
sequence; do not reuse an earlier value or call it concurrently from multiple
processes.

## Constraints

- Write rows in increasing order. Once a later row is written, earlier rows
  cannot be changed. Cells within the current row may be written in any column
  order.
- Configure column widths, hidden columns, freeze panes, autofilter, and tab
  color before writing data. `configure_sheet/2` rejects other instructions.
- This API writes a file and returns `:ok`. Returning the workbook as a binary
  would require memory proportional to its compressed XLSX size.
- The Rust writer keeps a temporary file for each sheet. Its default location
  is the system temporary directory. Set `tempdir:` to an existing disk-backed
  directory if the default is memory-backed or short on space.
- Peak memory depends on batch size, the widest row, and workbook metadata.
  Use a modest `batch_size` and avoid constructing all input rows beforehand.

For random access editing or a binary result, use `XlsxWriter.generate/2`.
