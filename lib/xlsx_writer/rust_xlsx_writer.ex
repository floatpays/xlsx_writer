defmodule XlsxWriter.RustXlsxWriter do
  @moduledoc false

  # quokka:skip-module-directives
  # `use RustlerPrecompiled` references the variables defined below, so it must
  # stay after them. The skip comment prevents Quokka from hoisting `use` to the
  # top of the module (which would break compilation).

  config = Mix.Project.config()

  version = config[:version]
  github_url = config[:package][:links]["GitHub"]

  # These targets and nif versions should correspond to the
  # .github/workflows/release.yml file.
  targets = ~w(
    aarch64-apple-darwin
    aarch64-unknown-linux-gnu
    aarch64-unknown-linux-musl
    riscv64gc-unknown-linux-gnu
    x86_64-apple-darwin
    x86_64-pc-windows-gnu
    x86_64-pc-windows-msvc
    x86_64-unknown-linux-gnu
    x86_64-unknown-linux-musl
  )

  nif_versions = ~w(
    2.16
  )

  use RustlerPrecompiled,
    otp_app: :xlsx_writer,
    crate: :xlsx_writer,
    base_url: "#{github_url}/releases/download/v#{version}",
    version: version,
    targets: targets,
    nif_versions: nif_versions

  def write(_data), do: :erlang.nif_error(:nif_not_loaded)

  def write_with_properties(_data, _properties),
    do: :erlang.nif_error(:nif_not_loaded)

  def stream_open(_tempdir, _properties), do: :erlang.nif_error(:nif_not_loaded)
  def stream_add_sheet(_resource, _name), do: :erlang.nif_error(:nif_not_loaded)

  def stream_configure_sheet(_resource, _instructions),
    do: :erlang.nif_error(:nif_not_loaded)

  def stream_write_cells(_resource, _cells),
    do: :erlang.nif_error(:nif_not_loaded)

  def stream_finish(_resource, _path), do: :erlang.nif_error(:nif_not_loaded)
  def stream_abort(_resource), do: :erlang.nif_error(:nif_not_loaded)
end
