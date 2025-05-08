defmodule XlsxWriter.RustXlsxWriter do
  use Rustler,
    otp_app: :xlsx_writer,
    crate: :rustxlsxwriter

  def test_new_workbook(), do: :erlang.nif_error(:nif_not_loaded)

  def get_binary(), do: :erlang.nif_error(:nif_not_loaded)

  def get_binary_with_image(_image_byte_list), do: :erlang.nif_error(:nif_not_loaded)
end
