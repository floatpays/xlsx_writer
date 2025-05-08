defmodule XlsxWriter do
  use Rustler,
    otp_app: :xlsx_writer,
    crate: :rustxlsxwriter

  def add(_a, _b), do: :erlang.nif_error(:nif_not_loaded)

  def test_new_workbook(), do: :erlang.nif_error(:nif_not_loaded)

  # To test that we can get a binary file back from rust code.
  def get_binary(), do: :erlang.nif_error(:nif_not_loaded)

  def test_binary() do
    {:ok, bytes} = get_binary() |> dbg
    IO.puts("Got bytes :)")
    IO.puts("Binary size: #{length(bytes)}")

    :ok = File.write("demo2.xlsx", bytes)
  end
end
