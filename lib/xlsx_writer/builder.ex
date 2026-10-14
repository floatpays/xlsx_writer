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

      XlsxWriter.Builder.create()
      |> XlsxWriter.Builder.add_sheet("Summary")
      |> XlsxWriter.Builder.add_rows([
        [{"Q1", format: [:bold]}, {"Q2", format: [:bold]}, {"Q3", format: [:bold]}],
        [100, 200, 300]
      ])
      |> XlsxWriter.Builder.skip_rows(1)
      |> XlsxWriter.Builder.add_rows([["Total", {600, format: [:italic]}]])
      |> XlsxWriter.Builder.write_file("output.xlsx")

  ## Features

  - Automatic cursor position tracking
  - Per-cell formatting via tuples
  - Multi-sheet support with `add_sheet/2`
  - Simple cursor movement with `skip_rows/2`
  - Seamless conversion to final XLSX binary

  ## Cell Format Options

  Each cell can be either a plain value or a tuple `{value, opts}` where opts include:

  - `width: number` - Column width (applies to entire column) - Builder-specific
  - `format: list` - List of XlsxWriter format options (see below)

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

  @type cell_value :: any()
  @type cell_with_format :: {cell_value(), keyword()}
  @type cell :: cell_value() | cell_with_format()
  @type row :: [cell()]

  @typedoc """
  Builder state.

  `sheets` holds the completed sheets, most recent first; the active sheet's
  instructions are kept in `instructions` until the next `add_sheet/3` or
  `write_binary/1`.
  """
  @type t :: %__MODULE__{
          sheets: [{String.t(), list()}],
          current_sheet: String.t() | nil,
          instructions: list(),
          cursor_row: non_neg_integer(),
          cursor_col: non_neg_integer(),
          column_widths: %{optional(non_neg_integer()) => number()}
        }

  defstruct sheets: [],
            current_sheet: nil,
            instructions: [],
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
  - `sheet_name` - Name of the new sheet
  - `opts` - Optional keyword list (reserved for future use)

  ## Returns

  Updated builder state with the new sheet active.

  ## Examples

      builder
      |> XlsxWriter.Builder.add_sheet("Summary")
      |> XlsxWriter.Builder.add_sheet("Details")

  """
  def add_sheet(%__MODULE__{} = builder, sheet_name, _opts \\ [])
      when is_binary(sheet_name) do
    builder = finalize_current_sheet(builder)

    %{
      builder
      | current_sheet: sheet_name,
        instructions: [],
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

  ## Examples

      # Simple rows
      builder
      |> XlsxWriter.Builder.add_rows([
        ["Name", "Age", "City"],
        ["Alice", 30, "NYC"],
        ["Bob", 25, "LA"]
      ])

      # With formatting
      builder
      |> XlsxWriter.Builder.add_rows([
        [{"Name", format: [:bold]}, {"Age", format: [:bold]}],
        ["Alice", {30, format: [{:num_format, "0"}]}]
      ])

      # With position override
      builder
      |> XlsxWriter.Builder.add_rows([["Data"]], start_row: 5, start_col: 2)

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

      builder
      |> XlsxWriter.Builder.add_rows([["Section 1"]])
      |> XlsxWriter.Builder.skip_rows(2)
      |> XlsxWriter.Builder.add_rows([["Section 2"]])

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

      {:ok, content} = builder |> XlsxWriter.Builder.write_binary()
      File.write!("output.xlsx", content)

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

  Convenience function that combines `write_binary/1` and `File.write!/2`.

  ## Parameters

  - `builder` - The builder state
  - `path` - The output file path (required)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      builder |> XlsxWriter.Builder.write_file("output.xlsx")
      builder |> XlsxWriter.Builder.write_file("reports/sales.xlsx")

  """
  def write_file(%__MODULE__{} = builder, path) when is_binary(path) do
    case write_binary(builder) do
      {:ok, content} ->
        File.write!(path, content)
        :ok

      {:error, reason} ->
        {:error, reason}
    end
  end

  # Private functions

  defp finalize_current_sheet(%__MODULE__{current_sheet: nil} = builder),
    do: builder

  # Moves the active sheet, with its column widths, into the completed sheets.
  defp finalize_current_sheet(%__MODULE__{} = builder) do
    width_instructions =
      Enum.map(builder.column_widths, fn {col, width} ->
        {:set_column_width, col, width}
      end)

    sheet =
      {builder.current_sheet, width_instructions ++ builder.instructions}

    %{
      builder
      | sheets: [sheet | builder.sheets],
        current_sheet: nil,
        instructions: [],
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

  defp write_cell(builder, row, col, {value, opts}) when is_list(opts) do
    # Extract Builder-specific options
    {width, remaining_opts} = Keyword.pop(opts, :width)
    {format, _remaining_opts} = Keyword.pop(remaining_opts, :format, [])

    builder =
      if width do
        put_in(builder.column_widths[col], width)
      else
        builder
      end

    # Pass format list directly to XlsxWriter
    add_write_instruction(builder, row, col, value, format)
  end

  defp write_cell(builder, row, col, value) do
    # Plain value without formatting
    add_write_instruction(builder, row, col, value, [])
  end

  # Prepends to the active sheet's instructions only, instead of rewriting
  # every sheet for every cell.
  defp add_write_instruction(builder, row, col, value, format) do
    {_name, instructions} =
      XlsxWriter.write(
        {builder.current_sheet, builder.instructions},
        row,
        col,
        value,
        format: format
      )

    %{builder | instructions: instructions}
  end
end
