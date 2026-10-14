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
  - [Builder API](guides/builder_api.md) - High-level API for quick data export (⚠️ Experimental)
  - [Advanced Formatting](guides/formatting.md) - Fonts, colors, borders, and number formats
  - [Layout Features](guides/layout_features.md) - Freeze panes, merged cells, autofilters, and more

  ## API Overview

  ### Core Functions
  - `generate/2` - Generate XLSX binary from sheets
  - `new_sheet/1` - Create a new worksheet

  ### Writing Data
  - `write/5` - Write any value to a cell
  - `write_formula/4` - Write Excel formula
  - `write_boolean/5` - Write boolean value
  - `write_url/5` - Write clickable URL
  - `write_image/4` - Embed image
  - `write_comment/5` - Add comment/note to cell
  - `write_blank/4` - Write formatted blank cell

  ### Layout & Structure
  - `set_column_width/3`, `set_row_height/3` - Size columns and rows
  - `set_column_range_width/4`, `set_row_range_height/4` - Size multiple columns/rows at once
  - `freeze_panes/3` - Lock rows/columns when scrolling
  - `merge_range/7` - Combine multiple cells
  - `hide_row/2`, `hide_column/2` - Hide rows/columns
  - `set_autofilter/5` - Add dropdown filters to headers

  See the [full documentation](https://hexdocs.pm/xlsx_writer) for detailed function references.
  """
  alias XlsxWriter.RustXlsxWriter
  alias XlsxWriter.Validation

  @doc """
  Generates an Excel xlsx file from a list of sheets.

  Takes a list of sheet tuples where each tuple contains a sheet name and
  a list of instructions for that sheet.

  ## Parameters

  - `sheets` - A list of `{sheet_name, instructions}` tuples
  - `opts` - Optional keyword list:
    - `:properties` - A `%XlsxWriter.WorkbookProperties{}` struct with document metadata

  ## Returns

  - `{:ok, xlsx_binary}` on success
  - `{:error, reason}` on failure, for example an invalid sheet name, a
    duplicate sheet name, an invalid URL or image, or an instruction that was
    not built with the functions in this module

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      ...>   |> XlsxWriter.write(0, 0, "Hello")
      iex> {:ok, xlsx_content} = XlsxWriter.generate([sheet])
      iex> is_binary(xlsx_content)
      true

      iex> sheet = XlsxWriter.new_sheet("Test") |> XlsxWriter.write(0, 0, "Hello")
      iex> properties = %XlsxWriter.WorkbookProperties{author: "John", title: "Report"}
      iex> {:ok, _content} = XlsxWriter.generate([sheet], properties: properties)

      iex> XlsxWriter.generate([XlsxWriter.new_sheet("A"), XlsxWriter.new_sheet("A")])
      {:error, "Worksheet name 'a' has already been used in this workbook."}

  """
  def generate(sheets, opts \\ []) when is_list(sheets) do
    sheets = Enum.map(sheets, &instructions_in_write_order/1)

    RustXlsxWriter.write(sheets, Keyword.get(opts, :properties))
  end

  # Instructions are prepended as they are added. Restore the order they were
  # added in, so that a later write to the same cell wins.
  defp instructions_in_write_order({name, instructions})
       when is_list(instructions), do: {name, Enum.reverse(instructions)}

  # Anything else is reported by the NIF as an invalid sheet.
  defp instructions_in_write_order(other), do: other

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

  Supports strings, numbers, booleans, `Decimal`, `Date`, `DateTime`,
  `NaiveDateTime`, atoms (written as strings) and `nil` (written as an empty
  string). Can also apply formatting options to the cell.

  `Date` values get the number format `yyyy-mm-dd` and `DateTime` /
  `NaiveDateTime` values `yyyy-mm-ddThh:mm:ss`, unless a `{:num_format, ...}`
  option overrides it.

  > #### Time zones {: .info}
  >
  > Excel has no notion of time zones, so a `DateTime` is written as its
  > wall-clock time and its UTC offset is dropped: `10:00:00+02:00` is
  > written as `10:00:00`. Shift it first (e.g. with `DateTime.shift_zone!/2`)
  > if you need the time in a particular zone.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index (0-based, at most 1,048,575)
  - `col` - The column index (0-based, at most 16,383)
  - `val` - The value to write
  - `opts` - Optional keyword list with formatting options

  ## Formatting Options

  - `:format` - A list of format specifications:
    - `:bold` - Make text bold
    - `:italic` - Make text italic
    - `:strikethrough` - Strike through text
    - `:superscript` - Superscript text
    - `:subscript` - Subscript text
    - `:text_wrap` - Wrap text within the cell
    - `{:rotation, angle}` - Rotate text (integer, -90 to 90 degrees, or 270 for vertical stacked text)
    - `:shrink` - Shrink text to fit within the cell width
    - `{:indent, level}` - Indent text by the given level (integer, 0 to 250)
    - `{:align, :left | :center | :right}` - Horizontal text alignment
    - `{:valign, :top | :center | :bottom | :justify | :distributed}` - Vertical text alignment
    - `{:num_format, format_string}` - Custom number format
    - `{:bg_color, hex_color}` - Background color (e.g., "#FFFF00" for yellow)
    - `{:font_color, hex_color}` - Font color (e.g., "#FF0000" for red)
    - `{:font_size, size}` - Font size in points, integer or float (e.g., 12, 10.5), greater than 0 and at most 409
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

  Colors are hex strings of the form `"#RRGGBB"` (case-insensitive; the `#`
  is optional).

  ## Border Styles

  Available border styles: `:thin`, `:medium`, `:thick`, `:dashed`, `:dotted`, `:double`,
  `:hair`, `:medium_dashed`, `:dash_dot`, `:medium_dash_dot`, `:dash_dot_dot`,
  `:medium_dash_dot_dot`, `:slant_dash_dot`

  ## Raises

  `ArgumentError` if the position is not a valid Excel cell, the value type
  is not supported, or a format option is unknown or invalid (see
  `XlsxWriter.Validation.validate_formats!/1`).

  ## Returns

  Updated sheet tuple with the new write instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write(sheet, 0, 0, "Hello")
      {"Test", [{:write, 0, 0, {:string, "Hello", []}}]}

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write(sheet, 0, 0, "Bold", format: [:bold])
      {"Test", [{:write, 0, 0, {:string, "Bold", [:bold]}}]}

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write(sheet, 0, 0, 42, format: [{:num_format, "0.00"}])
      {"Test", [{:write, 0, 0, {:number, 42.0, [{:num_format, "0.00"}]}}]}

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write(sheet, 0, 0, "Red Italic", format: [:italic, {:font_color, "#FF0000"}])
      {"Test", [{:write, 0, 0, {:string, "Red Italic", [:italic, {:font_color, "#FF0000"}]}}]}

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write(sheet, 0, 0, "Yellow", format: [{:bg_color, "yellow"}])
      ** (ArgumentError) Format option :bg_color expects a string hex color (e.g., "#FF0000"), got: "yellow"

  """
  def write({name, instructions}, row, col, val, opts \\ []) do
    Validation.validate_cell_position!(row, col)
    cell_data = to_cell_data(val, format_option!(opts))

    {name, [{:write, row, col, cell_data} | instructions]}
  end

  @doc """
  Writes an Excel formula to a specific cell in the sheet.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index (0-based)
  - `col` - The column index (0-based)
  - `val` - The Excel formula string (should start with '=')
  - `opts` - Optional keyword list with formatting options

  ## Options

  - `:format` - A list of format specifications (see `write/5`)

  ## Returns

  Updated sheet tuple with the new formula instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write_formula(sheet, 0, 2, "=A1+B1")
      {"Test", [{:write, 0, 2, {:formula, "=A1+B1", []}}]}

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write_formula(sheet, 0, 2, "=A1+B1", format: [:bold])
      {"Test", [{:write, 0, 2, {:formula, "=A1+B1", [:bold]}}]}

  """
  def write_formula({name, instructions}, row, col, val, opts \\ []) do
    Validation.validate_cell_position!(row, col)

    if !is_binary(val) do
      raise ArgumentError, "Formula must be a string, got: #{inspect(val)}"
    end

    formats = format_option!(opts)

    {name, [{:write, row, col, {:formula, val, formats}} | instructions]}
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
      iex> XlsxWriter.write_boolean(sheet, 0, 0, true)
      {"Test", [{:write, 0, 0, {:boolean, true, []}}]}

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write_boolean(sheet, 0, 0, false, format: [:bold])
      {"Test", [{:write, 0, 0, {:boolean, false, [:bold]}}]}

  """
  def write_boolean({name, instructions}, row, col, val, opts \\ [])
      when is_boolean(val) do
    Validation.validate_cell_position!(row, col)
    formats = format_option!(opts)

    {name, [{:write, row, col, {:boolean, val, formats}} | instructions]}
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
    - `:format` - Format specifications. Without it the cell gets Excel's
      default hyperlink style.

  ## Returns

  Updated sheet tuple with the new URL instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write_url(sheet, 0, 0, "https://example.com")
      {"Test", [{:write, 0, 0, {:url, "https://example.com", nil, []}}]}

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write_url(sheet, 0, 0, "https://example.com", text: "Click here")
      {"Test", [{:write, 0, 0, {:url, "https://example.com", "Click here", []}}]}

  """
  def write_url({name, instructions}, row, col, url, opts \\ [])
      when is_binary(url) do
    Validation.validate_cell_position!(row, col)
    text = Keyword.get(opts, :text)

    if !(is_binary(text) or is_nil(text)) do
      raise ArgumentError, "URL :text must be a string, got: #{inspect(text)}"
    end

    formats = format_option!(opts)

    {name, [{:write, row, col, {:url, url, text, formats}} | instructions]}
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
      iex> XlsxWriter.write_blank(sheet, 0, 0, format: [:bold, {:bg_color, "#FFFF00"}])
      {"Test", [{:write, 0, 0, {:blank, [:bold, {:bg_color, "#FFFF00"}]}}]}

  """
  def write_blank({name, instructions}, row, col, opts \\ []) do
    Validation.validate_cell_position!(row, col)
    formats = format_option!(opts)

    {name, [{:write, row, col, {:blank, formats}} | instructions]}
  end

  @doc """
  Writes a rich text string to a specific cell in the sheet.

  A rich string allows different formatting for different parts of the text
  within a single cell. Each segment consists of text and optional formatting.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index (0-based)
  - `col` - The column index (0-based)
  - `segments` - A list of `{text, formats}` tuples, where:
    - `text` is a non-empty string
    - `formats` is a list of format options (can be empty `[]` for default formatting)
  - `opts` - Optional keyword list with:
    - `:format` - Cell-level formatting (alignment, borders, background, etc.)

  ## Segment Format Options

  Each segment can have text formatting options:
  - `:bold` - Make text bold
  - `:italic` - Make text italic
  - `:strikethrough` - Strike through text
  - `:superscript` - Superscript text
  - `:subscript` - Subscript text
  - `{:font_color, hex_color}` - Font color (e.g., "#FF0000" for red)
  - `{:font_size, size}` - Font size in points
  - `{:font_name, name}` - Font family
  - `{:underline, style}` - Underline style

  ## Returns

  Updated sheet tuple with the new rich string instruction.

  ## Examples

      # Simple bold and normal text
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write_rich_string(sheet, 0, 0, [
      ...>   {"Bold ", [:bold]},
      ...>   {"Normal", []}
      ...> ])
      {"Test", [{:write, 0, 0, {:rich_string, [{"Bold ", [:bold]}, {"Normal", []}], []}}]}

      # Colored text segments
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write_rich_string(sheet, 0, 0, [
      ...>   {"Red ", [{:font_color, "#FF0000"}]},
      ...>   {"Blue", [{:font_color, "#0000FF"}]}
      ...> ])
      {"Test", [{:write, 0, 0, {:rich_string, [{"Red ", [{:font_color, "#FF0000"}]}, {"Blue", [{:font_color, "#0000FF"}]}], []}}]}

      # With cell-level formatting (centered)
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.write_rich_string(sheet, 0, 0, [
      ...>   {"Bold ", [:bold]},
      ...>   {"Italic", [:italic]}
      ...> ], format: [{:align, :center}])
      {"Test", [{:write, 0, 0, {:rich_string, [{"Bold ", [:bold]}, {"Italic", [:italic]}], [{:align, :center}]}}]}

  """
  def write_rich_string({name, instructions}, row, col, segments, opts \\ []) do
    Validation.validate_cell_position!(row, col)
    Validation.validate_rich_string_segments!(segments)
    formats = format_option!(opts)

    {name,
     [{:write, row, col, {:rich_string, segments, formats}} | instructions]}
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
    Validation.validate_cell_position!(row, col)
    Validation.validate_image_binary!(image_binary)

    {name, [{:write, row, col, {:image, image_binary}} | instructions]}
  end

  @doc """
  Writes a comment/note to a specific cell in the sheet.

  Comments appear when hovering over a cell and are useful for documentation,
  instructions, or additional context about cell values.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index (0-based)
  - `col` - The column index (0-based)
  - `text` - The comment text content
  - `opts` - Optional keyword list:
    - `:author` - Author name (string, max 52 characters)
    - `:visible` - Whether to show the comment by default (boolean, default: false)
    - `:width` - Comment box width in pixels (number >= 1, rounded to whole pixels, default: 128)
    - `:height` - Comment box height in pixels (number >= 1, rounded to whole pixels, default: 74)

  ## Returns

  Updated sheet tuple with the new comment instruction.

  ## Examples

      # Simple comment
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write_comment(sheet, 0, 0, "This is a note")
      iex> {"Test", [{:insert_note, 0, 0, "This is a note", _}]} = sheet

      # Comment with author
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write_comment(sheet, 0, 0, "Review this", author: "John Doe")
      iex> {"Test", [{:insert_note, 0, 0, "Review this", %XlsxWriter.NoteOptions{author: "John Doe"}}]} = sheet

      # Visible comment with custom size
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> sheet = XlsxWriter.write_comment(sheet, 0, 0, "Important!",
      ...>   visible: true, width: 300, height: 200)
      iex> {"Test", [{:insert_note, 0, 0, "Important!", options}]} = sheet
      iex> options.visible
      true
      iex> options.width
      300

  """
  def write_comment({name, instructions}, row, col, text, opts \\ []) do
    Validation.validate_cell_position!(row, col)

    if !is_binary(text) do
      raise ArgumentError,
            "Comment text must be a string, got: #{inspect(text)}"
    end

    Validation.validate_comment_options!(opts)

    note_options = %XlsxWriter.NoteOptions{
      author: Keyword.get(opts, :author),
      visible: Keyword.get(opts, :visible),
      width: Keyword.get(opts, :width),
      height: Keyword.get(opts, :height)
    }

    {name, [{:insert_note, row, col, text, note_options} | instructions]}
  end

  @doc """
  Sets the width of a specific column in the sheet.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `col` - The column index (0-based)
  - `width` - The width in Excel character units, an integer or float
    between 0 and 255

  ## Returns

  Updated sheet tuple with the new column width instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.set_column_width(sheet, 0, 25)
      {"Test", [{:set_column_width, 0, 25}]}

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.set_column_width(sheet, 0, 12.5)
      {"Test", [{:set_column_width, 0, 12.5}]}

  """
  def set_column_width({name, instructions}, col, width) do
    Validation.validate_column!(col)
    Validation.validate_column_width!(width)

    {name, [{:set_column_width, col, width} | instructions]}
  end

  @doc """
  Sets the height of a specific row in the sheet.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `row` - The row index (0-based)
  - `height` - The height in points, an integer or float between 0 and 409

  ## Returns

  Updated sheet tuple with the new row height instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.set_row_height(sheet, 0, 30.0)
      {"Test", [{:set_row_height, 0, 30.0}]}

  """
  def set_row_height({name, instructions}, row, height) do
    Validation.validate_row!(row)
    Validation.validate_row_height!(height)

    {name, [{:set_row_height, row, height} | instructions]}
  end

  @doc """
  Sets the width for a range of columns in the sheet.

  This is a convenience function to set the same width for multiple consecutive columns.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `first_col` - The first column index (0-based)
  - `last_col` - The last column index (0-based, inclusive)
  - `width` - The width value in Excel character units (matches `set_column_width/3`)

  ## Returns

  Updated sheet tuple with the new column range width instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.set_column_range_width(sheet, 0, 4, 20)
      {"Test", [{:set_column_range_width, 0, 4, 20}]}

  """
  def set_column_range_width({name, instructions}, first_col, last_col, width) do
    Validation.validate_column_range!(first_col, last_col)
    Validation.validate_column_width!(width)

    {name,
     [{:set_column_range_width, first_col, last_col, width} | instructions]}
  end

  @doc """
  Sets the height for a range of rows in the sheet.

  This is a convenience function to set the same height for multiple consecutive rows.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `first_row` - The first row index (0-based)
  - `last_row` - The last row index (0-based, inclusive)
  - `height` - The height value in points (matches `set_row_height/3`)

  ## Returns

  Updated sheet tuple with the new row range height instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.set_row_range_height(sheet, 0, 9, 25)
      {"Test", [{:set_row_range_height, 0, 9, 25}]}

  """
  def set_row_range_height({name, instructions}, first_row, last_row, height) do
    Validation.validate_row_range!(first_row, last_row)
    Validation.validate_row_height!(height)

    {name,
     [{:set_row_range_height, first_row, last_row, height} | instructions]}
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
      iex> XlsxWriter.freeze_panes(sheet, 1, 0)
      {"Test", [{:set_freeze_panes, 1, 0}]}

      # Freeze first column
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.freeze_panes(sheet, 0, 1)
      {"Test", [{:set_freeze_panes, 0, 1}]}

      # Freeze first row and first column
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.freeze_panes(sheet, 1, 1)
      {"Test", [{:set_freeze_panes, 1, 1}]}

  """
  def freeze_panes({name, instructions}, row, col) do
    Validation.validate_cell_position!(row, col)

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
      iex> XlsxWriter.hide_row(sheet, 5)
      {"Test", [{:set_row_hidden, 5}]}

  """
  def hide_row({name, instructions}, row) do
    Validation.validate_row!(row)

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
      iex> XlsxWriter.hide_column(sheet, 2)
      {"Test", [{:set_column_hidden, 2}]}

  """
  def hide_column({name, instructions}, col) do
    Validation.validate_column!(col)

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
      iex> XlsxWriter.set_autofilter(sheet, 0, 0, 0, 4)
      {"Test", [{:set_autofilter, 0, 0, 0, 4}]}

  """
  def set_autofilter(
        {name, instructions},
        first_row,
        first_col,
        last_row,
        last_col
      ) do
    Validation.validate_range!(first_row, first_col, last_row, last_col)

    {name,
     [
       {:set_autofilter, first_row, first_col, last_row, last_col}
       | instructions
     ]}
  end

  @doc """
  Enables autofit for all columns in the sheet.

  Automatically adjusts column widths to fit the longest content in each column.
  Autofit is applied after all data is written, wherever `autofit/1` appears
  in the pipeline, and explicit column widths set via `set_column_width/3`
  or `set_column_range_width/4` take precedence over it.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`

  ## Returns

  Updated sheet tuple with the autofit instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.autofit(sheet)
      {"Test", [:set_autofit]}

  """
  def autofit({name, instructions}) do
    {name, [:set_autofit | instructions]}
  end

  @doc """
  Sets the color of the worksheet tab.

  This changes the color of the sheet tab at the bottom of the Excel window,
  useful for visually organizing multi-sheet workbooks.

  ## Parameters

  - `sheet` - The sheet tuple `{name, instructions}`
  - `color` - Hex color string (e.g., "#FF0000" for red)

  ## Returns

  Updated sheet tuple with the tab color instruction.

  ## Examples

      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.set_tab_color(sheet, "#FF0000")
      {"Test", [{:set_tab_color, "#FF0000"}]}

  """
  def set_tab_color({name, instructions}, color) when is_binary(color) do
    Validation.validate_color!(color)

    {name, [{:set_tab_color, color} | instructions]}
  end

  @doc """
  Merges a range of cells into a single cell.

  The merged cell will contain the specified value and formatting.
  All merged cells will appear as one cell in Excel.

  The value can be any type `write/5` accepts and keeps its type: numbers,
  booleans and dates are written as such, not as text. Dates get the same
  default number formats as in `write/5`.

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
      iex> XlsxWriter.merge_range(sheet, 0, 0, 0, 3, "Title", format: [:bold, {:align, :center}])
      {"Test", [{:merge_range, 0, 0, 0, 3, {:string, "Title", [:bold, {:align, :center}]}}]}

      # Merge cells for a number
      iex> sheet = XlsxWriter.new_sheet("Test")
      iex> XlsxWriter.merge_range(sheet, 1, 1, 3, 1, 100)
      {"Test", [{:merge_range, 1, 1, 3, 1, {:number, 100.0, []}}]}

  """
  def merge_range(
        {name, instructions},
        first_row,
        first_col,
        last_row,
        last_col,
        val,
        opts \\ []
      ) do
    Validation.validate_range!(first_row, first_col, last_row, last_col)
    cell_data = to_cell_data(val, format_option!(opts))

    {name,
     [
       {:merge_range, first_row, first_col, last_row, last_col, cell_data}
       | instructions
     ]}
  end

  defp format_option!(opts) do
    formats = Keyword.get(opts, :format, [])
    Validation.validate_formats!(formats)
    formats
  end

  # Converts an Elixir value into the cell data understood by the NIF.
  defp to_cell_data(nil, formats), do: {:string, "", formats}

  defp to_cell_data(value, formats) when is_binary(value),
    do: {:string, value, formats}

  defp to_cell_data(value, formats) when is_boolean(value),
    do: {:boolean, value, formats}

  defp to_cell_data(value, formats) when is_atom(value),
    do: {:string, Atom.to_string(value), formats}

  defp to_cell_data(value, formats) when is_float(value),
    do: {:number, value, formats}

  defp to_cell_data(value, formats) when is_integer(value),
    do: {:number, integer_to_float!(value), formats}

  defp to_cell_data(%Decimal{} = value, formats),
    do: {:number, decimal_to_float!(value), formats}

  defp to_cell_data(%Date{} = value, formats),
    do: {:date, Date.to_iso8601(value), formats}

  # Excel has no time zones: write the wall-clock time and drop the offset.
  defp to_cell_data(%DateTime{} = value, formats) do
    {:date_time, value |> DateTime.to_naive() |> NaiveDateTime.to_iso8601(),
     formats}
  end

  defp to_cell_data(%NaiveDateTime{} = value, formats),
    do: {:date_time, NaiveDateTime.to_iso8601(value), formats}

  defp to_cell_data(value, _formats) do
    raise ArgumentError,
          "The data type for value #{inspect(value)} is not supported."
  end

  # Excel stores every number as a double.
  defp integer_to_float!(value) do
    :erlang.float(value)
  rescue
    ArgumentError ->
      reraise ArgumentError,
              [
                message:
                  "Integer #{value} is too large to be stored as an Excel number"
              ],
              __STACKTRACE__
  end

  defp decimal_to_float!(value) do
    Decimal.to_float(value)
  rescue
    # NaN, infinity, or beyond the range of a double.
    _error in [Decimal.Error, ArgumentError] ->
      reraise ArgumentError,
              [
                message:
                  "Decimal #{inspect(value)} cannot be stored as an Excel number"
              ],
              __STACKTRACE__
  end
end
