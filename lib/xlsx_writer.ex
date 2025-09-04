defmodule XlsxWriter do
  @moduledoc """
  A library for creating Excel xlsx files in Elixir.

  XlsxWriter provides a simple API for generating Excel spreadsheets with support for
  various data types, formatting, formulas, images, and layout customization.

  ## Basic Usage

  Create a new sheet and write data to it:

      iex> sheet = XlsxWriter.new_sheet("My Sheet")
      iex> sheet = XlsxWriter.write(sheet, 0, 0, "Hello")
      iex> sheet = XlsxWriter.write(sheet, 0, 1, "World")
      iex> {:ok, _xlsx_content} = XlsxWriter.generate([sheet])

  ## Supported Data Types

  XlsxWriter automatically handles various Elixir data types:

      iex> sheet = XlsxWriter.new_sheet("Data Types")
      iex> sheet = sheet
      ...>   |> XlsxWriter.write(0, 0, "String")
      ...>   |> XlsxWriter.write(1, 0, 42)
      ...>   |> XlsxWriter.write(2, 0, 3.14)
      ...>   |> XlsxWriter.write(3, 0, Date.utc_today())
      ...>   |> XlsxWriter.write(4, 0, DateTime.utc_now())
      ...>   |> XlsxWriter.write(5, 0, Decimal.new("99.99"))
      iex> {:ok, _xlsx_content} = XlsxWriter.generate([sheet])

  ## Formatting

  Apply formatting to cells using the `:format` option:

      iex> sheet = XlsxWriter.new_sheet("Formatted")
      iex> sheet = sheet
      ...>   |> XlsxWriter.write(0, 0, "Bold Text", format: [:bold])
      ...>   |> XlsxWriter.write(0, 1, "Centered", format: [{:align, :center}])
      ...>   |> XlsxWriter.write(0, 2, 1234.56)
      iex> {:ok, _xlsx_content} = XlsxWriter.generate([sheet])

  ## Formulas

  Write Excel formulas to cells:

      iex> sheet = XlsxWriter.new_sheet("Formulas")
      iex> sheet = sheet
      ...>   |> XlsxWriter.write(0, 0, 10)
      ...>   |> XlsxWriter.write(0, 1, 20)
      ...>   |> XlsxWriter.write_formula(0, 2, "=A1+B1")
      iex> {:ok, _xlsx_content} = XlsxWriter.generate([sheet])

  ## Column and Row Sizing

  Customize column widths and row heights:

      iex> sheet = XlsxWriter.new_sheet("Sized")
      iex> sheet = sheet
      ...>   |> XlsxWriter.write(0, 0, "Wide Column")
      ...>   |> XlsxWriter.set_column_width(0, 25)
      ...>   |> XlsxWriter.set_row_height(0, 40)
      iex> {:ok, _xlsx_content} = XlsxWriter.generate([sheet])

  ## Multiple Sheets

  Create workbooks with multiple sheets:

      iex> sheet1 = XlsxWriter.new_sheet("First Sheet")
      ...>   |> XlsxWriter.write(0, 0, "Sheet 1 Data")
      iex> sheet2 = XlsxWriter.new_sheet("Second Sheet")
      ...>   |> XlsxWriter.write(0, 0, "Sheet 2 Data")
      iex> {:ok, _xlsx_content} = XlsxWriter.generate([sheet1, sheet2])

  ## Complete Example

  Here's a comprehensive example showing various features:

      iex> sheet = XlsxWriter.new_sheet("Sales Report")
      ...>   |> XlsxWriter.write(0, 0, "Product", format: [:bold])
      ...>   |> XlsxWriter.write(0, 1, "Quantity", format: [:bold])
      ...>   |> XlsxWriter.write(0, 2, "Price", format: [:bold])
      ...>   |> XlsxWriter.write(0, 3, "Total", format: [:bold])
      ...>   |> XlsxWriter.write(1, 0, "Widget A")
      ...>   |> XlsxWriter.write(1, 1, 100)
      ...>   |> XlsxWriter.write(1, 2, 9.99)
      ...>   |> XlsxWriter.write_formula(1, 3, "=B2*C2")
      ...>   |> XlsxWriter.set_column_width(0, 15)
      ...>   |> XlsxWriter.set_column_width(1, 12)
      ...>   |> XlsxWriter.set_column_width(2, 12)
      ...>   |> XlsxWriter.set_column_width(3, 12)
      iex> {:ok, _xlsx_content} = XlsxWriter.generate([sheet])
  """
  alias XlsxWriter.RustXlsxWriter

  @doc """
  Generates an Excel xlsx file from a list of sheets.

  Takes a list of sheet tuples where each tuple contains a sheet name and
  a list of instructions for that sheet.

  ## Parameters

  - `sheets` - A list of `{sheet_name, instructions}` tuples

  ## Returns

  - `{:ok, xlsx_binary}` on success
  - `{:error, reason}` on failure

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      ...>   |> XlsxWriter.write(0, 0, "Hello")
      iex> {:ok, xlsx_content} = XlsxWriter.generate([sheet])
      iex> is_binary(xlsx_content)
      true

  """
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

  @doc """
  Creates a new empty sheet with the given name.

  ## Parameters

  - `name` - The name of the sheet (must be a string)

  ## Returns

  A sheet tuple `{name, []}` ready for writing data.

  ## Examples

      iex> XlsxWriter.new_sheet("My Sheet")
      {"My Sheet", []}

  """
  def new_sheet(name) when is_binary(name), do: {name, []}

  @doc """
  Writes a value to a specific cell in the sheet.

  Supports various data types including strings, numbers, dates, and Decimal values.
  Can also apply formatting options to the cell.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index (0-based)
  - `col` - The column index (0-based)
  - `val` - The value to write
  - `opts` - Optional keyword list with formatting options

  ## Formatting Options

  - `:format` - A list of format specifications:
    - `:bold` - Make text bold
    - `{:align, :left | :center | :right}` - Text alignment
    - `{:num_format, format_string}` - Custom number format

  ## Returns

  Updated sheet tuple with the new write instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write(sheet, 0, 0, "Hello")
      iex> {"Test", [{:write, 0, 0, {:string, "Hello"}}]} = sheet

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write(sheet, 0, 0, "Bold", format: [:bold])
      iex> {"Test", [{:write, 0, 0, {:string_with_format, "Bold", [:bold]}}]} = sheet

  """
  def write({name, instructions}, row, col, val, opts \\ []) do
    case Keyword.get(opts, :format) do
      nil ->
        {name, [{:write, row, col, to_rust_val(val)} | instructions]}
      
      formats when is_list(formats) ->
        write_with_format({name, instructions}, row, col, val, formats)
    end
  end

  @doc """
  Writes an Excel formula to a specific cell in the sheet.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index (0-based)
  - `col` - The column index (0-based)
  - `val` - The Excel formula string (should start with '=')

  ## Returns

  Updated sheet tuple with the new formula instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write_formula(sheet, 0, 2, "=A1+B1")
      iex> {"Test", [{:write, 0, 2, {:formula, "=A1+B1"}}]} = sheet

  """
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

  @doc """
  Writes an image to a specific cell in the sheet.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index (0-based)
  - `col` - The column index (0-based)
  - `image_binary` - The binary content of the image file

  ## Returns

  Updated sheet tuple with the new image instruction.

  ## Examples

      iex> image_data = <<137, 80, 78, 71>>  # Mock PNG header
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write_image(sheet, 0, 0, image_data)
      iex> {"Test", [{:write, 0, 0, {:image, ^image_data}}]} = sheet

  """
  def write_image({name, instructions}, row, col, image_binary) do
    {name, [{:write, row, col, {:image, image_binary}} | instructions]}
  end

  @doc """
  Sets the width of a specific column in the sheet.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `col` - The column index (0-based)
  - `width` - The width value (typically a float)

  ## Returns

  Updated sheet tuple with the new column width instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.set_column_width(sheet, 0, 25)
      iex> {"Test", [{:set_column_width, 0, 25}]} = sheet

  """
  def set_column_width({name, instructions}, col, width) do
    {name, [{:set_column_width, col, width} | instructions]}
  end

  @doc """
  Sets the height of a specific row in the sheet.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index (0-based)
  - `height` - The height value (typically a float)

  ## Returns

  Updated sheet tuple with the new row height instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.set_row_height(sheet, 0, 30.0)
      iex> {"Test", [{:set_row_height, 0, 30.0}]} = sheet

  """
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
