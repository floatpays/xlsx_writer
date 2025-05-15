# XlsxWriter

<!-- MDOC !-->

Xlsx Writer

## Usage

## Installation

If [available in Hex](https://hex.pm/docs/publish), the package can be installed
by adding `xlsx_writer` to your list of dependencies in `mix.exs`:

```elixir
def deps do
  [
    {:xlsx_writer, "~> 0.1.3"}
  ]
end
```

Documentation can be generated with [ExDoc](https://github.com/elixir-lang/ex_doc)
and published on [HexDocs](https://hexdocs.pm). Once published, the docs can
be found at <https://hexdocs.pm/xlsx_writer>.

## Development

### Publishing a new version

As per instruction: https://hexdocs.pm/rustler_precompiled/precompilation_guide.html

- release a new tag
- push the code to your repository with the new tag: git push origin main --tags
- wait for all NIFs to be built
- run the mix rustler_precompiled.download task (with the flag --all)
- release the package to Hex.pm (make sure your release includes the correct files).


    mix rustler_precompiled.download XlsxWriter.RustXlsxWriter --all


## Copyright and License

Copyright (c) 2025 Floatpays

This work is free. You can redistribute it and/or modify it under the
terms of the MIT License. See the [LICENSE.md](./LICENSE.md) file for more details.
