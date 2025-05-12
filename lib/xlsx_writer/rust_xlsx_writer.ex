defmodule XlsxWriter.RustXlsxWriter do
  @moduledoc false

  use RustlerPrecompiled,
    otp_app: :xlsx_writer,
    crate: :rustxlsxwriter,
    base_url:
      "https://github.com/floatpays/xlsx_writer/releases/download/v0.2.0",
    version: "0.2.0",
    nif_versions: ["2.17"]

  def write(_data), do: :erlang.nif_error(:nif_not_loaded)
end
