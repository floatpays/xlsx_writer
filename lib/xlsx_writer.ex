defmodule XlsxWriter do
  alias XlsxWriter.RustXlsxWriter

  def write(data, filename) do
    with {:ok, bytes} <- RustXlsxWriter.write(data),
         :ok <- File.write(filename, bytes) do
      {:ok, filename}
    end
  end
end
