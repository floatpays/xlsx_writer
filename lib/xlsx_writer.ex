defmodule XlsxWriter.Workbook do
  @moduledoc false
  alias XlsxWriter.RustXlsxWriter

  def generate(sheets) when is_list(sheets) do
    # It might not be important to reverse the instructions here
    # but doing it to avoid potential confusion.
    sheets =
      Enum.map(sheets, fn {name, instructions} ->
        {name, Enum.reverse(instructions)}
      end)

    case RustXlsxWriter.write(sheets) do
      {:ok, content} ->
        {:ok, IO.iodata_to_binary(content)}

      other ->
        other
    end
  end

  def new_sheet(name) when is_binary(name), do: {name, []}

  def write({name, instructions}, row, col, val, opts \\ []) do
    case Keyword.get(opts, :format) do
      nil ->
        {name, [{:write, row, col, to_rust_val(val)} | instructions]}
      
      formats when is_list(formats) ->
        write_with_format({name, instructions}, row, col, val, formats)
    end
  end

  def write_formula({name, instructions}, row, col, val) do
    {name, [{:write, row, col, {:formula, val}} | instructions]}
  end

  defp write_with_format({name, instructions}, row, col, val, formats)
       when is_binary(val) do
    instruction = {:write, row, col, {:string_with_format, val, formats}}

    {name, [instruction | instructions]}
  end

  defp write_with_format({name, instructions}, row, col, numeric_val, formats)
       when is_number(numeric_val) do
    instruction =
      {:write, row, col, {:number_with_format, numeric_val, formats}}

    {name, [instruction | instructions]}
  end

  def write_image({name, instructions}, row, col, image_binary) do
    {name, [{:write, row, col, {:image, image_binary}} | instructions]}
  end

  def set_column_width({name, instructions}, col, width) do
    {name, [{:set_column_width, col, width} | instructions]}
  end

  def set_row_height({name, instructions}, row, height) do
    {name, [{:set_row_height, row, height} | instructions]}
  end

  defp to_rust_val(val) do
    case val do
      %Decimal{} = amount ->
        {:float, Decimal.to_float(amount)}

      %Date{} = date ->
        {:date, Date.to_iso8601(date)}

      %DateTime{} = datetime ->
        {:date_time, DateTime.to_iso8601(datetime)}

      %NaiveDateTime{} = datetime ->
        {:date_time, NaiveDateTime.to_iso8601(datetime)}

      val when is_binary(val) ->
        {:string, val}

      val when is_float(val) ->
        {:float, val}

      val when is_integer(val) ->
        {:float, val}

      val when is_nil(val) ->
        {:string, ""}

      val when is_atom(val) ->
        {:string, Atom.to_string(val)}

      other ->
        raise XlsxWriter.Error,
              "The data type for value \"#{inspect(other)}\" is not supported."
    end
  end
end
