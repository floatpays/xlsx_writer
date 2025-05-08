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

  def get_binary_with_image(_image_byte_list), do: :erlang.nif_error(:nif_not_loaded)

  # To show our rust wrapper code can get an image as bytes, and return the workbook as bytes.
  def test_binary_with_image() do
    image_byte_list = File.read!("bird.jpeg") |> :binary.bin_to_list()

    {:ok, bytes} = get_binary_with_image(image_byte_list)

    IO.puts("Got bytes, with image :)")
    IO.puts("Binary size: #{length(bytes)}")

    :ok = File.write("demo3.xlsx", bytes)
  end
end
