defmodule XlsxWriter.Builder do
  @moduledoc """
  A high-level builder API for creating Excel files without manually tracking cell positions.

  > #### Experimental Feature {: .warning}
  >
  > This API is experimental and subject to change in future releases. While functional
  > and tested, the API design may evolve based on user feedback. Use with caution in
  > production code and expect potential breaking changes.

  This module provides a convenient way to generate Excel files by automatically
  tracking cursor positions as you add data. It's ideal for simple use cases where
  you're generating large files without complex layout requirements.

  ## Quick Example

      iex> {:ok, content} =
      ...>   XlsxWriter.Builder.create()
      ...>   |> XlsxWriter.Builder.add_sheet("Summary")
      ...>   |> XlsxWriter.Builder.add_rows([
      ...>     [{"Q1", format: [:bold]}, {"Q2", format: [:bold]}, {"Q3", format: [:bold]}],
      ...>     [100, 200, 300]
      ...>   ])
      ...>   |> XlsxWriter.Builder.skip_rows(1)
      ...>   |> XlsxWriter.Builder.add_rows([["Total", {600, format: [:italic]}]])
      ...>   |> XlsxWriter.Builder.write_binary()
      iex> binary_part(content, 0, 2)
      "PK"

  Use `write_file/2` instead of `write_binary/1` to write the workbook to disk.

  ## Features

  - Automatic cursor position tracking
  - Per-cell formatting via tuples
  - Multi-sheet support with `add_sheet/2`
  - Simple cursor movement with `skip_rows/2`
  - Seamless conversion to final XLSX binary

  ## Cell Format Options

  Each cell can be either a plain value or a tuple `{value, opts}` where opts include:

  - `width: number` - Column width in character units, integer or float
    (applies to entire column) - Builder-specific
  - `format: list` - List of XlsxWriter format options (see below)

  Any other option raises `ArgumentError`.

  The `format` option accepts the same format list as `XlsxWriter.write/5`:

  - `:bold`, `:italic`, `:strikethrough` - Text styles
  - `{:font_size, number}` - Font size in points
  - `{:font_color, "#RRGGBB"}` - Font color (hex)
  - `{:bg_color, "#RRGGBB"}` - Background color (hex)
  - `{:align, :left | :center | :right}` - Text alignment
  - `{:num_format, "format_string"}` - Number format
  - `{:border, style}`, `{:border_top, style}`, etc. - Borders

  See `XlsxWriter.write/5` for complete format options.
  """

  alias XlsxWriter
  alias XlsxWriter.Validation

  @cell_options [:format, :width]

  @type cell_value :: any()
  @type cell_with_format :: {cell_value(), keyword()}
  @type cell :: cell_value() | cell_with_format()
  @type row :: [cell()]

  @typedoc """
  Builder state.

  `sheets` holds the completed sheets, most recent first; the active sheet
  (named `current_sheet`) is kept in `sheet` until the next `add_sheet/3` or
  `write_binary/1`. Sheets are `t:XlsxWriter.sheet/0` values.
  """
  @type t :: %__MODULE__{
          sheets: [XlsxWriter.sheet()],
          current_sheet: String.t() | nil,
          sheet: XlsxWriter.sheet() | nil,
          cursor_row: non_neg_integer(),
          cursor_col: non_neg_integer(),
          column_widths: %{optional(non_neg_integer()) => number()}
        }

  defstruct sheets: [],
            current_sheet: nil,
            sheet: nil,
            cursor_row: 0,
            cursor_col: 0,
            column_widths: %{}

  @doc """
  Creates a new builder for generating an Excel file.

  ## Returns

  A new builder state.

  ## Examples

      iex> builder = XlsxWriter.Builder.create()
      iex> is_struct(builder, XlsxWriter.Builder)
      true

  """
  def create do
    %__MODULE__{}
  end

  @doc """
  Adds a new sheet to the workbook and switches context to it.

  The cursor position resets to (0, 0) for the new sheet.

  ## Parameters

  - `builder` - The builder state
  - `sheet_name` - Name of the new sheet. Like Excel, sheet names must be
    unique ignoring case; a duplicate makes `write_binary/1` return an error.
  - `opts` - Optional keyword list (reserved for future use)

  ## Returns

  Updated builder state with the new sheet active.

  ## Examples

      iex> builder =
      ...>   XlsxWriter.Builder.create()
      ...>   |> XlsxWriter.Builder.add_sheet("Summary")
      ...>   |> XlsxWriter.Builder.add_sheet("Details")
      iex> builder.current_sheet
      "Details"

  """
  def add_sheet(%__MODULE__{} = builder, sheet_name, _opts \\ [])
      when is_binary(sheet_name) do
    builder = finalize_current_sheet(builder)

    %{
      builder
      | current_sheet: sheet_name,
        sheet: XlsxWriter.new_sheet(sheet_name),
        cursor_row: 0,
        cursor_col: 0,
        column_widths: %{}
    }
  end

  @doc """
  Adds multiple rows starting at the current cursor position.

  Each row is a list of cells. Cells can be plain values or tuples `{value, opts}`
  for formatting.

  ## Parameters

  - `builder` - The builder state
  - `rows` - List of rows, where each row is a list of cells
  - `opts` - Optional keyword list:
    - `:start_row` - Override cursor row position (0-based)
    - `:start_col` - Override cursor column position (0-based)

  ## Returns

  Updated builder state with cursor moved after the last row.

  ## Raises

  `ArgumentError` if no sheet was added, a cell has an unknown option, or a
  value, position or format is invalid (see `XlsxWriter.write/5`).

  ## Examples

      iex> builder =
      ...>   XlsxWriter.Builder.create()
      ...>   |> XlsxWriter.Builder.add_sheet("People")
      ...>   |> XlsxWriter.Builder.add_rows([
      ...>     [{"Name", format: [:bold]}, {"Age", format: [:bold]}],
      ...>     ["Alice", {30, format: [{:num_format, "0"}]}],
      ...>     ["Bob", 25]
      ...>   ])
      iex> {builder.cursor_row, builder.cursor_col}
      {3, 0}

      iex> builder =
      ...>   XlsxWriter.Builder.create()
      ...>   |> XlsxWriter.Builder.add_sheet("Positioned")
      ...>   |> XlsxWriter.Builder.add_rows([["Data"]], start_row: 5, start_col: 2)
      iex> {builder.cursor_row, builder.cursor_col}
      {6, 2}

      iex> XlsxWriter.Builder.create()
      ...> |> XlsxWriter.Builder.add_sheet("Typo")
      ...> |> XlsxWriter.Builder.add_rows([[{"x", formats: [:bold]}]])
      ** (ArgumentError) Unknown cell option :formats in {"x", [formats: [:bold]]}, expected one of [:format, :width]

  """
  def add_rows(builder, rows, opts \\ [])

  def add_rows(%__MODULE__{current_sheet: nil}, _rows, _opts) do
    raise ArgumentError, "No active sheet. Call add_sheet/2 first."
  end

  def add_rows(%__MODULE__{} = builder, rows, opts) when is_list(rows) do
    start_row = Keyword.get(opts, :start_row, builder.cursor_row)
    start_col = Keyword.get(opts, :start_col, builder.cursor_col)

    {builder, final_row} =
      Enum.reduce(rows, {builder, start_row}, fn row, {builder, row_index} ->
        {add_single_row(builder, row, row_index, start_col), row_index + 1}
      end)

    %{builder | cursor_row: final_row, cursor_col: start_col}
  end

  @doc """
  Moves the cursor down by N rows.

  Useful for adding spacing between sections of data.

  ## Parameters

  - `builder` - The builder state
  - `n` - Number of rows to skip (default: 1)

  ## Returns

  Updated builder state with cursor moved down.

  ## Examples

      iex> builder =
      ...>   XlsxWriter.Builder.create()
      ...>   |> XlsxWriter.Builder.add_sheet("Sections")
      ...>   |> XlsxWriter.Builder.add_rows([["Section 1"]])
      ...>   |> XlsxWriter.Builder.skip_rows(2)
      ...>   |> XlsxWriter.Builder.add_rows([["Section 2"]])
      iex> builder.cursor_row
      4

  """
  def skip_rows(%__MODULE__{} = builder, n \\ 1) when is_integer(n) and n > 0 do
    %{builder | cursor_row: builder.cursor_row + n}
  end

  @doc """
  Generates the final Excel file and returns the binary content.

  This finalizes all sheets and generates the XLSX file binary.

  ## Parameters

  - `builder` - The builder state

  ## Returns

  - `{:ok, xlsx_binary}` on success
  - `{:error, reason}` on failure

  ## Examples

      iex> {:ok, content} =
      ...>   XlsxWriter.Builder.create()
      ...>   |> XlsxWriter.Builder.add_sheet("Data")
      ...>   |> XlsxWriter.Builder.add_rows([["Hello", "World"]])
      ...>   |> XlsxWriter.Builder.write_binary()
      iex> is_binary(content)
      true

      iex> XlsxWriter.Builder.create() |> XlsxWriter.Builder.write_binary()
      {:error, "No sheets added. Use add_sheet/2 to add at least one sheet."}

  """
  def write_binary(%__MODULE__{} = builder) do
    case finalize_current_sheet(builder) do
      %__MODULE__{sheets: []} ->
        {:error, "No sheets added. Use add_sheet/2 to add at least one sheet."}

      %__MODULE__{sheets: sheets} ->
        sheets
        |> Enum.reverse()
        |> XlsxWriter.generate()
    end
  end

  @doc """
  Generates the final Excel file and writes it to disk.

  Convenience function that combines `write_binary/1` and `File.write/2`.

  ## Parameters

  - `builder` - The builder state
  - `path` - The output file path (required)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure: the reason from `write_binary/1`, or a
    `File.write/2` reason such as `:enoent` if the file cannot be written

  ## Examples

      builder |> XlsxWriter.Builder.write_file("output.xlsx")
      #=> :ok

      builder |> XlsxWriter.Builder.write_file("missing_directory/sales.xlsx")
      #=> {:error, :enoent}

  """
  def write_file(%__MODULE__{} = builder, path) when is_binary(path) do
    with {:ok, content} <- write_binary(builder) do
      File.write(path, content)
    end
  end

  # Private functions

  # Moves the active sheet, with its column widths, into the completed sheets.
  defp finalize_current_sheet(%__MODULE__{current_sheet: nil} = builder),
    do: builder

  # Explicit widths take precedence over autofit wherever they are set.
  defp finalize_current_sheet(%__MODULE__{} = builder) do
    sheet =
      Enum.reduce(builder.column_widths, builder.sheet, fn {col, width},
                                                           sheet ->
        XlsxWriter.set_column_width(sheet, col, width)
      end)

    %{
      builder
      | sheets: [sheet | builder.sheets],
        current_sheet: nil,
        sheet: nil,
        column_widths: %{}
    }
  end

  defp add_single_row(builder, row, row_index, start_col) do
    row
    |> Enum.with_index(start_col)
    |> Enum.reduce(builder, fn {cell, col_index}, builder ->
      write_cell(builder, row_index, col_index, cell)
    end)
  end

  defp write_cell(builder, row, col, {value, opts} = cell) when is_list(opts) do
    validate_cell_options!(cell, opts)

    builder =
      case Keyword.get(opts, :width) do
        nil ->
          builder

        width ->
          Validation.validate_column!(col)
          Validation.validate_column_width!(width)
          put_in(builder.column_widths[col], width)
      end

    add_write_instruction(
      builder,
      row,
      col,
      value,
      Keyword.get(opts, :format) || []
    )
  end

  defp write_cell(builder, row, col, value) do
    add_write_instruction(builder, row, col, value, [])
  end

  defp validate_cell_options!(cell, opts) do
    if !Keyword.keyword?(opts) do
      raise ArgumentError,
            "Cell options must be a keyword list, got: #{inspect(cell)}"
    end

    case Keyword.keys(opts) -- @cell_options do
      [] ->
        :ok

      [unknown_option | _] ->
        raise ArgumentError,
              "Unknown cell option #{inspect(unknown_option)} in #{inspect(cell)}, expected one of #{inspect(@cell_options)}"
    end
  end

  defp add_write_instruction(builder, row, col, value, format) do
    sheet = XlsxWriter.write(builder.sheet, row, col, value, format: format)

    %{builder | sheet: sheet}
  end
end
