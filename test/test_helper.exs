defmodule XlsxWriter.TestHelpers do
  @moduledoc """
  Helpers for asserting on the contents of generated workbooks.
  """

  @doc """
  Unzips an xlsx binary in memory and returns the XML of `part_name`
  (e.g. `"xl/worksheets/sheet1.xml"`, `"xl/styles.xml"`,
  `"xl/sharedStrings.xml"`), or `nil` when the part does not exist.
  """
  def xlsx_part(xlsx_binary, part_name) when is_binary(xlsx_binary) do
    {:ok, files} = :zip.unzip(xlsx_binary, [:memory])
    part_charlist = String.to_charlist(part_name)

    Enum.find_value(files, fn
      {^part_charlist, content} -> content
      _other -> nil
    end)
  end

  @doc """
  Generates a workbook from `sheets` and returns the XML of `part_name`.
  """
  def generate_part!(sheets, part_name \\ "xl/worksheets/sheet1.xml") do
    {:ok, xlsx_binary} = XlsxWriter.generate(sheets)
    xlsx_part(xlsx_binary, part_name)
  end
end

ExUnit.start()
