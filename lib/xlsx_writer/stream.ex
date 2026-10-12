defmodule XlsxWriter.Stream do
  @moduledoc """
  Writes large workbooks to a file with bounded input batches.

  Rows must be written in increasing order. A row may contain plain values or
  `{value, format: formats}` cells, using the same values and formats as
  `XlsxWriter.write/5`. Formulas may be `{:formula, expression}` or
  `{:formula, expression, format: formats}`. The workbook is saved directly to
  disk; no XLSX binary is returned. Each sheet uses `rust_xlsxwriter`'s
  constant memory mode.

  For an Enumerable of rows, `write_file/3` batches input automatically:

      rows = Elixir.Stream.map(1..100_000, fn n -> [n, "Row \#{n}"] end)
      XlsxWriter.Stream.write_file("large.xlsx", [{"Data", rows}])

  The lower-level API accepts bounded lists of rows:

      {:ok, writer} = XlsxWriter.Stream.open("large.xlsx")
      {:ok, writer} = XlsxWriter.Stream.add_sheet(writer, "Data")
      {:ok, writer} = XlsxWriter.Stream.write_rows(writer, [["Name"], ["Alice"]])
      :ok = XlsxWriter.Stream.finish(writer)

  Call `abort/1` if stopping before `finish/1` to release the workbook's
  temporary files promptly. `write_file/3` does this automatically.
  """

  alias XlsxWriter.RustXlsxWriter

  defstruct [:resource, :path, next_row: 0]

  @type t :: %__MODULE__{
          resource: reference(),
          path: String.t(),
          next_row: non_neg_integer()
        }

  @doc """
  Opens a workbook for writing to `path`.

  Options are `:tempdir` for worksheet temporary files and `:properties` for
  `%XlsxWriter.WorkbookProperties{}`. The temporary directory must already exist.
  """
  def open(path, opts \\ []) when is_binary(path) and is_list(opts) do
    tempdir = Keyword.get(opts, :tempdir)
    properties = Keyword.get(opts, :properties)

    with :ok <- validate_options(tempdir, properties),
         {:ok, resource} <- RustXlsxWriter.stream_open(tempdir, properties) do
      {:ok, %__MODULE__{resource: resource, path: path}}
    end
  end

  @doc "Adds a constant memory worksheet and makes it active."
  def add_sheet(%__MODULE__{} = writer, name) when is_binary(name) do
    case RustXlsxWriter.stream_add_sheet(writer.resource, name) do
      {:ok, {}} -> {:ok, %{writer | next_row: 0}}
      error -> error
    end
  end

  @doc """
  Configures the active sheet before writing rows.

  Build an instruction tuple with `XlsxWriter.new_sheet/1`. Supported operations
  are column widths, hidden columns, freeze panes, autofilter, and tab color.
  The tuple's sheet name is ignored because `add_sheet/2` already set it.
  """
  def configure_sheet(%__MODULE__{} = writer, {_name, instructions})
      when is_list(instructions) do
    writer.resource
    |> RustXlsxWriter.stream_configure_sheet(Enum.reverse(instructions))
    |> case do
      {:ok, {}} -> {:ok, writer}
      error -> error
    end
  end

  @doc """
  Writes a bounded list of rows to the active sheet.

  `:start_row` can skip forward. `:start_col` sets the first column for all
  rows in this call. Both indices are zero based. A nonempty call advances the
  cursor to the row after its last row.
  """
  def write_rows(%__MODULE__{} = writer, rows, opts \\ [])
      when is_list(rows) and is_list(opts) do
    start_row = Keyword.get(opts, :start_row, writer.next_row)
    start_col = Keyword.get(opts, :start_col, 0)

    cond do
      not (is_integer(start_row) and start_row >= writer.next_row and
               start_row <= 1_048_575) ->
        {:error, "start_row must be between the next row and 1,048,575"}

      not (is_integer(start_col) and start_col >= 0 and start_col <= 16_383) ->
        {:error, "start_col must be between 0 and 16,383"}

      start_row + length(rows) > 1_048_576 ->
        {:error, "Rows exceed the Excel worksheet limit"}

      true ->
        cells =
          rows
          |> Enum.with_index(start_row)
          |> Enum.flat_map(fn {row, row_index} ->
            row
            |> Enum.with_index(start_col)
            |> Enum.map(fn {cell, col_index} ->
              data = cell_to_data(cell, row_index, col_index)
              {row_index, col_index, data}
            end)
          end)

        case RustXlsxWriter.stream_write_cells(writer.resource, cells) do
          {:ok, {}} -> {:ok, %{writer | next_row: start_row + length(rows)}}
          error -> error
        end
    end
  end

  @doc "Finishes the workbook and writes the XLSX file to the path passed to `open/2`."
  def finish(%__MODULE__{} = writer) do
    normalize_nif_result(
      RustXlsxWriter.stream_finish(writer.resource, writer.path)
    )
  end

  @doc "Releases an unfinished workbook and its temporary files."
  def abort(%__MODULE__{} = writer) do
    normalize_nif_result(RustXlsxWriter.stream_abort(writer.resource))
  end

  @doc """
  Writes an Enumerable of `{sheet_name, rows}` pairs to `path`.

  Each `rows` value must be an Enumerable of row lists. `:batch_size` defaults
  to 500; the remaining options are passed to `open/2`.
  """
  def write_file(path, sheets, opts \\ [])
      when is_binary(path) and is_list(opts) do
    batch_size = Keyword.get(opts, :batch_size, 500)

    if is_integer(batch_size) and batch_size > 0 do
      with {:ok, writer} <- open(path, opts) do
        try do
          result =
            Enum.reduce_while(sheets, {:ok, writer}, fn {name, rows},
                                                        {:ok, current} ->
              case add_sheet(current, name) do
                {:ok, current} ->
                  rows
                  |> Elixir.Stream.chunk_every(batch_size)
                  |> Enum.reduce_while({:ok, current}, fn batch,
                                                          {:ok, writer} ->
                    case write_rows(writer, batch) do
                      {:ok, writer} -> {:cont, {:ok, writer}}
                      error -> {:halt, error}
                    end
                  end)
                  |> case do
                    {:ok, writer} -> {:cont, {:ok, writer}}
                    error -> {:halt, error}
                  end

                error ->
                  {:halt, error}
              end
            end)

          case result do
            {:ok, writer} -> finish(writer)
            error -> error
          end
        after
          abort(writer)
        end
      end
    else
      {:error, "batch_size must be a positive integer"}
    end
  end

  defp cell_to_data({:formula, expression}, row, col),
    do: cell_to_data({:formula, expression, []}, row, col)

  defp cell_to_data({:formula, expression, opts}, row, col)
       when is_binary(expression) and is_list(opts) do
    {_, [{:write, ^row, ^col, data}]} =
      XlsxWriter.write_formula({"", []}, row, col, expression, opts)

    data
  end

  defp cell_to_data({value, opts}, row, col) when is_list(opts) do
    {_, [{:write, ^row, ^col, data}]} =
      XlsxWriter.write({"", []}, row, col, value, opts)

    data
  end

  defp cell_to_data(value, row, col) do
    {_, [{:write, ^row, ^col, data}]} =
      XlsxWriter.write({"", []}, row, col, value)

    data
  end

  defp normalize_nif_result({:ok, {}}), do: :ok
  defp normalize_nif_result(error), do: error

  defp validate_options(tempdir, properties) do
    cond do
      not (is_nil(tempdir) or is_binary(tempdir)) ->
        {:error, "tempdir must be a path string"}

      not (is_nil(properties) or
               is_struct(properties, XlsxWriter.WorkbookProperties)) ->
        {:error, "properties must be a XlsxWriter.WorkbookProperties struct"}

      true ->
        :ok
    end
  end
end
