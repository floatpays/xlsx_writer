defmodule XlsxWriter do
  @moduledoc """
  A high-performance library for creating Excel xlsx files in Elixir.

  Built with the powerful `rust_xlsxwriter` crate via Rustler NIF for excellent speed
  and memory efficiency. Supports rich formatting, formulas, images, and advanced layout features.

  ## Quick Start

      sheet = XlsxWriter.new_sheet("My Sheet")
        |> XlsxWriter.write(0, 0, "Hello", format: [:bold])
        |> XlsxWriter.write(0, 1, "World")

      {:ok, xlsx_content} = XlsxWriter.generate([sheet])
      File.write!("output.xlsx", xlsx_content)

  ## Key Features

  - **Data Types**: Strings, numbers, dates, booleans, URLs, formulas, images
  - **Rich Formatting**: Fonts, colors, borders, alignment, number formats
  - **Layout Control**: Freeze panes, merged cells, autofilters, hide rows/columns
  - **High Performance**: Rust-powered NIF for fast generation of large spreadsheets

  ## Guides

  - [Getting Started](guides/getting_started.md) - Basic usage and data types
  - [Advanced Formatting](guides/formatting.md) - Fonts, colors, borders, and number formats
  - [Layout Features](guides/layout_features.md) - Freeze panes, merged cells, autofilters, and more

  ## API Overview

  ### Core Functions
  - `generate/1` - Generate XLSX binary from sheets
  - `new_sheet/1` - Create a new worksheet

  ### Writing Data
  - `write/5` - Write any value to a cell
  - `write_formula/4` - Write Excel formula
  - `write_boolean/5` - Write boolean value
  - `write_url/5` - Write clickable URL
  - `write_image/4` - Embed image
  - `write_blank/4` - Write formatted blank cell

  ### Layout & Structure
  - `set_column_width/3`, `set_row_height/3` - Size columns and rows
  - `freeze_panes/3` - Lock rows/columns when scrolling
  - `merge_range/7` - Combine multiple cells
  - `hide_row/2`, `hide_column/2` - Hide rows/columns
  - `set_autofilter/5` - Add dropdown filters to headers

  See the [full documentation](https://hexdocs.pm/xlsx_writer) for detailed function references.
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
    - `:italic` - Make text italic
    - `:strikethrough` - Strike through text
    - `:superscript` - Superscript text
    - `:subscript` - Subscript text
    - `{:align, :left | :center | :right}` - Text alignment
    - `{:num_format, format_string}` - Custom number format
    - `{:bg_color, hex_color}` - Background color (e.g., "#FFFF00" for yellow)
    - `{:font_color, hex_color}` - Font color (e.g., "#FF0000" for red)
    - `{:font_size, size}` - Font size in points (e.g., 12, 14, 16)
    - `{:font_name, name}` - Font family (e.g., "Arial", "Times New Roman")
    - `{:underline, :single | :double | :single_accounting | :double_accounting}` - Underline style
    - `{:pattern, :solid | :none | :gray125 | :gray0625}` - Fill pattern
    - `{:border, style}` - Apply border to all sides (see border styles below)
    - `{:border_top, style}` - Top border
    - `{:border_bottom, style}` - Bottom border
    - `{:border_left, style}` - Left border
    - `{:border_right, style}` - Right border
    - `{:border_color, hex_color}` - Color for all borders
    - `{:border_top_color, hex_color}` - Top border color
    - `{:border_bottom_color, hex_color}` - Bottom border color
    - `{:border_left_color, hex_color}` - Left border color
    - `{:border_right_color, hex_color}` - Right border color

  ## Border Styles

  Available border styles: `:thin`, `:medium`, `:thick`, `:dashed`, `:dotted`, `:double`,
  `:hair`, `:medium_dashed`, `:dash_dot`, `:medium_dash_dot`, `:dash_dot_dot`,
  `:medium_dash_dot_dot`, `:slant_dash_dot`

  ## Returns

  Updated sheet tuple with the new write instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write(sheet, 0, 0, "Hello")
      iex> {"Test", [{:write, 0, 0, {:string, "Hello"}}]} = sheet

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write(sheet, 0, 0, "Bold", format: [:bold])
      iex> {"Test", [{:write, 0, 0, {:string_with_format, "Bold", [:bold]}}]} = sheet

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write(sheet, 0, 0, "Yellow", format: [{:bg_color, "#FFFF00"}])
      iex> {"Test", [{:write, 0, 0, {:string_with_format, "Yellow", [{:bg_color, "#FFFF00"}]}}]} = sheet

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write(sheet, 0, 0, "Red Italic", format: [:italic, {:font_color, "#FF0000"}])
      iex> {"Test", [{:write, 0, 0, {:string_with_format, "Red Italic", [:italic, {:font_color, "#FF0000"}]}}]} = sheet

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write(sheet, 0, 0, "Bordered", format: [{:border, :thin}])
      iex> {"Test", [{:write, 0, 0, {:string_with_format, "Bordered", [{:border, :thin}]}}]} = sheet

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

  @doc """
  Writes a boolean value to a specific cell in the sheet.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index (0-based)
  - `col` - The column index (0-based)
  - `val` - The boolean value (true or false)
  - `opts` - Optional keyword list with formatting options

  ## Returns

  Updated sheet tuple with the new boolean instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write_boolean(sheet, 0, 0, true)
      iex> {"Test", [{:write, 0, 0, {:boolean, true}}]} = sheet

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write_boolean(sheet, 0, 0, false, format: [:bold])
      iex> {"Test", [{:write, 0, 0, {:boolean_with_format, false, [:bold]}}]} = sheet

  """
  def write_boolean({name, instructions}, row, col, val, opts \\ []) when is_boolean(val) do
    case Keyword.get(opts, :format) do
      nil ->
        {name, [{:write, row, col, {:boolean, val}} | instructions]}

      formats when is_list(formats) ->
        {name, [{:write, row, col, {:boolean_with_format, val, formats}} | instructions]}
    end
  end

  @doc """
  Writes a URL/hyperlink to a specific cell in the sheet.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index (0-based)
  - `col` - The column index (0-based)
  - `url` - The URL string
  - `opts` - Optional keyword list with:
    - `:text` - Display text (different from URL)
    - `:format` - Format specifications

  ## Returns

  Updated sheet tuple with the new URL instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write_url(sheet, 0, 0, "https://example.com")
      iex> {"Test", [{:write, 0, 0, {:url, "https://example.com"}}]} = sheet

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write_url(sheet, 0, 0, "https://example.com", text: "Click here")
      iex> {"Test", [{:write, 0, 0, {:url_with_text, "https://example.com", "Click here"}}]} = sheet

  """
  def write_url({name, instructions}, row, col, url, opts \\ []) when is_binary(url) do
    text = Keyword.get(opts, :text)
    formats = Keyword.get(opts, :format)

    instruction =
      case {text, formats} do
        {nil, nil} ->
          {:write, row, col, {:url, url}}

        {text, nil} when is_binary(text) ->
          {:write, row, col, {:url_with_text, url, text}}

        {nil, formats} when is_list(formats) ->
          {:write, row, col, {:url_with_format, url, formats}}

        {text, formats} when is_binary(text) and is_list(formats) ->
          {:write, row, col, {:url_with_text_and_format, url, text, formats}}
      end

    {name, [instruction | instructions]}
  end

  @doc """
  Writes a blank cell with formatting to the sheet.

  A blank cell differs from an empty cell - it has no data but can have formatting.
  This is useful for pre-formatting cells before data is added.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index (0-based)
  - `col` - The column index (0-based)
  - `opts` - Keyword list with `:format` specifications

  ## Returns

  Updated sheet tuple with the new blank cell instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write_blank(sheet, 0, 0, format: [:bold, {:bg_color, "#FFFF00"}])
      iex> {"Test", [{:write, 0, 0, {:blank, [:bold, {:bg_color, "#FFFF00"}]}}]} = sheet

  """
  def write_blank({name, instructions}, row, col, opts \\ []) do
    formats = Keyword.get(opts, :format, [])
    {name, [{:write, row, col, {:blank, formats}} | instructions]}
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

  @doc """
  Freezes panes at the specified row and column.

  This locks rows and/or columns so they remain visible when scrolling.
  Very useful for keeping headers visible.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row to freeze at (0-based). Rows above this remain visible.
  - `col` - The column to freeze at (0-based). Columns left of this remain visible.

  ## Returns

  Updated sheet tuple with the freeze panes instruction.

  ## Examples

      # Freeze the first row (header row)
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.freeze_panes(sheet, 1, 0)
      iex> {"Test", [{:set_freeze_panes, 1, 0}]} = sheet

      # Freeze first column
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.freeze_panes(sheet, 0, 1)
      iex> {"Test", [{:set_freeze_panes, 0, 1}]} = sheet

      # Freeze first row and first column
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.freeze_panes(sheet, 1, 1)
      iex> {"Test", [{:set_freeze_panes, 1, 1}]} = sheet

  """
  def freeze_panes({name, instructions}, row, col) do
    {name, [{:set_freeze_panes, row, col} | instructions]}
  end

  @doc """
  Hides a specific row in the sheet.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index to hide (0-based)

  ## Returns

  Updated sheet tuple with the hide row instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.hide_row(sheet, 5)
      iex> {"Test", [{:set_row_hidden, 5}]} = sheet

  """
  def hide_row({name, instructions}, row) do
    {name, [{:set_row_hidden, row} | instructions]}
  end

  @doc """
  Hides a specific column in the sheet.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `col` - The column index to hide (0-based)

  ## Returns

  Updated sheet tuple with the hide column instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.hide_column(sheet, 2)
      iex> {"Test", [{:set_column_hidden, 2}]} = sheet

  """
  def hide_column({name, instructions}, col) do
    {name, [{:set_column_hidden, col} | instructions]}
  end

  @doc """
  Sets an autofilter on a range of cells.

  Adds dropdown filter buttons to the specified range, typically used on header rows.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `first_row` - The first row of the filter range (0-based)
  - `first_col` - The first column of the filter range (0-based)
  - `last_row` - The last row of the filter range (0-based)
  - `last_col` - The last column of the filter range (0-based)

  ## Returns

  Updated sheet tuple with the autofilter instruction.

  ## Examples

      # Set autofilter on header row (row 0, columns A-E)
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.set_autofilter(sheet, 0, 0, 0, 4)
      iex> {"Test", [{:set_autofilter, 0, 0, 0, 4}]} = sheet

  """
  def set_autofilter({name, instructions}, first_row, first_col, last_row, last_col) do
    {name, [{:set_autofilter, first_row, first_col, last_row, last_col} | instructions]}
  end

  @doc """
  Merges a range of cells into a single cell.

  The merged cell will contain the specified value and formatting.
  All merged cells will appear as one cell in Excel.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `first_row` - The first row of the merge range (0-based)
  - `first_col` - The first column of the merge range (0-based)
  - `last_row` - The last row of the merge range (0-based)
  - `last_col` - The last column of the merge range (0-based)
  - `val` - The value to write in the merged cell
  - `opts` - Optional keyword list with formatting options

  ## Returns

  Updated sheet tuple with the merge range instruction.

  ## Examples

      # Merge cells A1:D1 with centered title
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.merge_range(sheet, 0, 0, 0, 3, "Title", format: [:bold, {:align, :center}])
      iex> {"Test", [{:merge_range, 0, 0, 0, 3, {:string_with_format, "Title", [:bold, {:align, :center}]}}]} = sheet

      # Merge cells for a number
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.merge_range(sheet, 1, 1, 3, 1, 100)
      iex> {"Test", [{:merge_range, 1, 1, 3, 1, {:float, 100}}]} = sheet

  """
  def merge_range({name, instructions}, first_row, first_col, last_row, last_col, val, opts \\ []) do
    case Keyword.get(opts, :format) do
      nil ->
        {name, [{:merge_range, first_row, first_col, last_row, last_col, to_rust_val(val)} | instructions]}

      formats when is_list(formats) ->
        merge_range_with_format(
          {name, instructions},
          first_row,
          first_col,
          last_row,
          last_col,
          val,
          formats
        )
    end
  end

  defp merge_range_with_format(
         {name, instructions},
         first_row,
         first_col,
         last_row,
         last_col,
         val,
         formats
       )
       when is_binary(val) do
    instruction =
      {:merge_range, first_row, first_col, last_row, last_col,
       {:string_with_format, val, formats}}

    {name, [instruction | instructions]}
  end

  defp merge_range_with_format(
         {name, instructions},
         first_row,
         first_col,
         last_row,
         last_col,
         numeric_val,
         formats
       )
       when is_number(numeric_val) do
    instruction =
      {:merge_range, first_row, first_col, last_row, last_col,
       {:number_with_format, numeric_val, formats}}

    {name, [instruction | instructions]}
  end

  defp merge_range_with_format(
         {name, instructions},
         first_row,
         first_col,
         last_row,
         last_col,
         val,
         formats
       )
       when is_boolean(val) do
    instruction =
      {:merge_range, first_row, first_col, last_row, last_col,
       {:boolean_with_format, val, formats}}

    {name, [instruction | instructions]}
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
