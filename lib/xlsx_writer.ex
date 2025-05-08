defmodule XlsxWriter do
  alias XlsxWriter.RustXlsxWriter

  def test_binary() do
    {:ok, bytes} = RustXlsxWriter.get_binary() |> dbg
    IO.puts("Got bytes :)")
    IO.puts("Binary size: #{length(bytes)}")

    :ok = File.write("demo2.xlsx", bytes)
  end

  # To show our rust wrapper code can get an image as bytes, and return the workbook as bytes.
  def test_binary_with_image() do
    image_byte_list = File.read!("bird.jpeg") |> :binary.bin_to_list()

    {:ok, bytes} = RustXlsxWriter.get_binary_with_image(image_byte_list)

    IO.puts("Got bytes, with image :)")
    IO.puts("Binary size: #{length(bytes)}")

    :ok = File.write("demo3.xlsx", bytes)
  end
end
