# XlsxWriter

<!-- MDOC !-->

Xlsx Writer

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

If [available in Hex](https://hex.pm/docs/publish), the package can be installed
by adding `xlsx_writer` to your list of dependencies in `mix.exs`:

```elixir
def deps do
  [
    {:xlsx_writer, "~> 0.1.4"}
  ]
end
```

Documentation can be generated with [ExDoc](https://github.com/elixir-lang/ex_doc)
and published on [HexDocs](https://hexdocs.pm). Once published, the docs can
be found at <https://hexdocs.pm/xlsx_writer>.

## Development

### Publishing a new version

As per instruction: https://hexdocs.pm/rustler_precompiled/precompilation_guide.html

- update version number in mix.exs, and in this file.
- release a new tag for the new version, eg. "v0.1.4"
- push the code to your repository with the new tag: git push origin main --tags
  - wait for github actions to complete successfully.
    (then all NIFs will be built)
- run the mix rustler_precompiled.download task (with the flag --all)
  `mix rustler_precompiled.download XlsxWriter.RustXlsxWriter --all`

- release the package to Hex.pm (make sure your release includes the correct files).
  `mix hex.publish`



## Copyright and License

Copyright (c) 2025 Floatpays

This work is free. You can redistribute it and/or modify it under the
terms of the MIT License. See the [LICENSE.md](./LICENSE.md) file for more details.
