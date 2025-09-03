# XlsxWriter

<!-- MDOC !-->

A fast Elixir library for writing Excel files using Rust.

## Usage

```elixir
filename = "test2.xlsx"

sheet1 =
  Workbook.new_sheet("sheet number one")
  |> Workbook.write_with_format(0, 0, "col1", [:bold])
  |> Workbook.write_with_format(0, 1, "col2", [:bold, {:align, :center}])
  |> Workbook.write_with_format(0, 2, "col3", [:bold, {:align, :right}])
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

## Installation

Add `xlsx_writer` to your list of dependencies in `mix.exs`:

```elixir
def deps do
  [
    {:xlsx_writer, "~> 0.1.4"}
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
