defmodule XlsxWriter.RustXlsxWriter do
  @moduledoc false

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
end
