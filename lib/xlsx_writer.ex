defmodule XlsxWriter do
  use Rustler,
    otp_app: :xlsx_writer,
    crate: :rustxlsxwriter

  def add(_a, _b), do: :erlang.nif_error(:nif_not_loaded)

  def test_new_workbook(), do: :erlang.nif_error(:nif_not_loaded)

  def new_workbook(), do: :erlang.nif_error(:nif_not_loaded)

  def write_to_workbook(), do: :erlang.nif_error(:nif_not_loaded)
end
