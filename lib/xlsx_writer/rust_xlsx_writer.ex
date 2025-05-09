defmodule XlsxWriter.RustXlsxWriter do
  use Rustler,
    otp_app: :xlsx_writer,
    crate: :rustxlsxwriter

  def write(_data), do: :erlang.nif_error(:nif_not_loaded)
end
