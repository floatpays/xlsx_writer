defmodule XlsxWriter.BuilderTest do
  use ExUnit.Case, async: true

  import XlsxWriter.TestHelpers

  alias XlsxWriter.Builder

  doctest XlsxWriter.Builder

  describe "create/0" do
    test "creates a new builder" do
      builder = Builder.create()
      assert builder.sheets == []
      assert builder.current_sheet == nil
      assert builder.cursor_row == 0
      assert builder.cursor_col == 0
    end
  end

  describe "add_sheet/2" do
    test "adds a new sheet and switches context to it" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")

      assert builder.current_sheet == "Sheet1"
      assert builder.cursor_row == 0
      assert builder.cursor_col == 0
      assert builder.sheets == []
    end

    test "supports multiple sheets" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_sheet("Sheet2")

      assert builder.current_sheet == "Sheet2"
      assert [{"Sheet1", []}] = builder.sheets

      {:ok, content} = Builder.write_binary(builder)
      assert xlsx_part(content, "xl/worksheets/sheet1.xml")
      assert xlsx_part(content, "xl/worksheets/sheet2.xml")
      refute xlsx_part(content, "xl/worksheets/sheet3.xml")
    end

    test "raises when a sheet name is already used, ignoring case" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Data")
        |> Builder.add_rows([["first"]])
        |> Builder.add_sheet("Other")

      for name <- ["Data", "DATA", "Other"] do
        assert_raise ArgumentError,
                     "A sheet named #{inspect(name)} already exists (sheet names are case-insensitive)",
                     fn -> Builder.add_sheet(builder, name) end
      end
    end

    test "rows only go to the active sheet" do
      {:ok, content} =
        Builder.create()
        |> Builder.add_sheet("First")
        |> Builder.add_rows([["first"]])
        |> Builder.add_sheet("Second")
        |> Builder.add_rows([["second"]])
        |> Builder.write_binary()

      shared_strings = xlsx_part(content, "xl/sharedStrings.xml")

      assert xlsx_part(content, "xl/worksheets/sheet1.xml") =~
               ~s(<c r="A1" t="s"><v>#{shared_string_index(shared_strings, "first")}</v>)

      assert xlsx_part(content, "xl/worksheets/sheet2.xml") =~
               ~s(<c r="A1" t="s"><v>#{shared_string_index(shared_strings, "second")}</v>)
    end

    test "resets cursor when switching sheets" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([["A", "B"]])
        |> Builder.add_sheet("Sheet2")

      assert builder.cursor_row == 0
      assert builder.cursor_col == 0
    end
  end

  describe "add_rows/2" do
    # Passed through the test context so the type checker cannot prove the
    # call below fails, which is exactly what this test asserts at runtime.
    setup do
      %{builder_without_sheet: Builder.create()}
    end

    test "raises error when no active sheet", %{builder_without_sheet: builder} do
      assert_raise ArgumentError,
                   "No active sheet. Call add_sheet/2 first.",
                   fn ->
                     Builder.add_rows(builder, [["A", "B"]])
                   end
    end

    test "adds simple rows and advances cursor" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([
          ["A", "B", "C"],
          [1, 2, 3]
        ])

      assert builder.cursor_row == 2
      assert builder.cursor_col == 0
    end

    test "handles rows with formatting tuples" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([
          [{"Header 1", format: [:bold]}, {"Header 2", format: [:bold]}],
          ["Value 1", {42, format: [{:num_format, "0.00"}]}]
        ])

      assert builder.cursor_row == 2
    end

    test "respects start_row and start_col options" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([["Data"]], start_row: 5, start_col: 3)

      assert builder.cursor_row == 6
      assert builder.cursor_col == 3
    end

    test "cursor position is updated after using start_row and start_col" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([["First"]])
        # Cursor is now at (1, 0)
        |> Builder.add_rows([["Override"]], start_row: 5, start_col: 3)

      # Cursor should now be at (6, 3) - one row after the override position

      assert builder.cursor_row == 6
      assert builder.cursor_col == 3

      # Adding more rows should continue from the new cursor position
      builder = Builder.add_rows(builder, [["Next"]])
      assert builder.cursor_row == 7
      assert builder.cursor_col == 3
    end

    test "start_row and start_col can be used independently" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([["A"]])
        # Cursor at (1, 0)
        |> Builder.add_rows([["B"]], start_row: 5)

      # Only row overridden, col stays at 0

      assert builder.cursor_row == 6
      assert builder.cursor_col == 0

      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([["A"]])
        # Cursor at (1, 0)
        |> Builder.add_rows([["B"]], start_col: 5)

      # Only col overridden, row continues from cursor

      assert builder.cursor_row == 2
      assert builder.cursor_col == 5
    end

    test "handles column widths in format options" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([[{"Wide Column", width: 30}]])

      assert Map.has_key?(builder.column_widths, 0)
      assert builder.column_widths[0] == 30
    end

    test "accepts fractional column widths" do
      {:ok, content} =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([[{"x", width: 12.5}, {"y", width: 20}]])
        |> Builder.write_binary()

      sheet_xml = xlsx_part(content, "xl/worksheets/sheet1.xml")

      assert sheet_xml =~
               ~s(<col min="1" max="1" width="13.28515625" customWidth="1"/>)

      assert sheet_xml =~
               ~s(<col min="2" max="2" width="20.7109375" customWidth="1"/>)
    end

    test "raises for an invalid column width" do
      for width <- [-1, 256, "wide"] do
        assert_raise ArgumentError, ~r/Column width must be a number/, fn ->
          Builder.create()
          |> Builder.add_sheet("Sheet1")
          |> Builder.add_rows([[{"x", width: width}]])
        end
      end
    end

    test "format: nil and width: nil are the same as leaving them out" do
      write = fn rows ->
        {:ok, content} =
          Builder.create()
          |> Builder.add_sheet("Sheet1")
          |> Builder.add_rows(rows)
          |> Builder.write_binary()

        {xlsx_part(content, "xl/worksheets/sheet1.xml"),
         xlsx_part(content, "xl/styles.xml")}
      end

      assert write.([[{"x", format: nil}, {"y", width: nil}]]) ==
               write.([["x", "y"]])
    end

    test "raises for unknown cell options instead of ignoring them" do
      assert_raise ArgumentError,
                   ~s(Unknown cell option :formats in {"x", [formats: [:bold]]}, expected one of [:format, :width]),
                   fn ->
                     Builder.create()
                     |> Builder.add_sheet("Sheet1")
                     |> Builder.add_rows([[{"x", formats: [:bold]}]])
                   end
    end

    test "raises for cell options that are not a keyword list" do
      assert_raise ArgumentError, ~r/Cell options must be a keyword list/, fn ->
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([[{"x", [:bold]}]])
      end
    end

    test "raises for invalid formats at call time" do
      assert_raise ArgumentError,
                   ~r/:bg_color expects a string hex color/,
                   fn ->
                     Builder.create()
                     |> Builder.add_sheet("Sheet1")
                     |> Builder.add_rows([
                       [{"x", format: [{:bg_color, "red"}]}]
                     ])
                   end
    end

    test "writes multiple cells with different formats" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([
          [
            {"Bold", format: [:bold]},
            {"Italic", format: [:italic]},
            {"Both", format: [:bold, :italic]}
          ]
        ])

      assert builder.cursor_row == 1
    end
  end

  describe "skip_rows/2" do
    test "moves cursor down by specified number of rows" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([["Row 1"]])
        |> Builder.skip_rows(3)

      assert builder.cursor_row == 4
    end

    test "defaults to skipping 1 row" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([["Row 1"]])
        |> Builder.skip_rows()

      assert builder.cursor_row == 2
    end

    test "can be chained multiple times" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.skip_rows(2)
        |> Builder.skip_rows(3)

      assert builder.cursor_row == 5
    end
  end

  describe "write_binary/1" do
    test "returns error when no sheets added" do
      builder = Builder.create()

      assert {:error, reason} = Builder.write_binary(builder)
      assert reason =~ "No sheets added"
    end

    test "generates valid XLSX binary" do
      {:ok, content} =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([["Hello", "World"]])
        |> Builder.write_binary()

      assert is_binary(content)
      # XLSX files start with PK (ZIP magic bytes)
      assert binary_part(content, 0, 2) == "PK"
    end

    test "handles multiple sheets" do
      {:ok, content} =
        Builder.create()
        |> Builder.add_sheet("Summary")
        |> Builder.add_rows([["Total", 100]])
        |> Builder.add_sheet("Details")
        |> Builder.add_rows([["Item", "Amount"], ["A", 50], ["B", 50]])
        |> Builder.write_binary()

      assert is_binary(content)
      assert binary_part(content, 0, 2) == "PK"
    end
  end

  describe "write_file/2" do
    @tag :tmp_dir
    test "returns an error tuple when the file cannot be written", %{
      tmp_dir: dir
    } do
      path = Path.join([dir, "missing_directory", "output.xlsx"])

      result =
        Builder.create()
        |> Builder.add_sheet("Test")
        |> Builder.write_file(path)

      assert result == {:error, :enoent}
    end

    @tag :tmp_dir
    test "returns the generation error without writing a file", %{tmp_dir: dir} do
      path = Path.join(dir, "output.xlsx")

      assert {:error, "No sheets added" <> _} =
               Builder.write_file(Builder.create(), path)

      refute File.exists?(path)
    end

    @tag :tmp_dir
    test "writes XLSX file to disk", %{tmp_dir: dir} do
      path = Path.join(dir, "output.xlsx")

      result =
        Builder.create()
        |> Builder.add_sheet("Test")
        |> Builder.add_rows([["Data"]])
        |> Builder.write_file(path)

      assert result == :ok
      assert File.exists?(path)
    end

    @tag :tmp_dir
    test "writes to specified path", %{tmp_dir: dir} do
      path = Path.join(dir, "custom.xlsx")

      result =
        Builder.create()
        |> Builder.add_sheet("Test")
        |> Builder.add_rows([["Data"]])
        |> Builder.write_file(path)

      assert result == :ok
      assert File.exists?(path)
    end
  end

  describe "integration tests" do
    test "complex example with formatting and multiple sheets" do
      {:ok, content} =
        Builder.create()
        |> Builder.add_sheet("Summary")
        |> Builder.add_rows([
          [
            {"Q1", format: [:bold]},
            {"Q2", format: [:bold]},
            {"Q3", format: [:bold]}
          ],
          [100, 200, 300]
        ])
        |> Builder.skip_rows(1)
        |> Builder.add_rows([
          [
            {"Total", format: [:italic]},
            {600, format: [{:num_format, "$#,##0.00"}]}
          ]
        ])
        |> Builder.add_sheet("Details")
        |> Builder.add_rows([
          [
            {"Name", format: [:bold], width: 20},
            {"Value", format: [:bold], width: 15}
          ],
          ["Item A", 100],
          ["Item B", 200],
          ["Item C", 300]
        ])
        |> Builder.write_binary()

      assert is_binary(content)
    end

    test "large data set" do
      data = Enum.map(1..100, fn i -> ["Row #{i}", i, i * 2] end)

      {:ok, content} =
        Builder.create()
        |> Builder.add_sheet("Data")
        |> Builder.add_rows([["Col1", "Col2", "Col3"] | data])
        |> Builder.write_binary()

      assert is_binary(content)
    end

    test "mixed data types" do
      {:ok, content} =
        Builder.create()
        |> Builder.add_sheet("Mixed")
        |> Builder.add_rows([
          ["String", 123, 45.67, true],
          ["Another", 456, 78.90, false]
        ])
        |> Builder.write_binary()

      assert is_binary(content)
    end

    test "using start_row and start_col to position data" do
      {:ok, content} =
        Builder.create()
        |> Builder.add_sheet("Data")
        |> Builder.add_rows([["Top Left"]], start_row: 0, start_col: 0)
        |> Builder.add_rows([["Offset"]], start_row: 5, start_col: 5)
        |> Builder.write_binary()

      assert is_binary(content)
    end

    test "column widths are applied correctly" do
      builder =
        Builder.create()
        |> Builder.add_sheet("Sheet1")
        |> Builder.add_rows([[{"Col1", width: 20}, {"Col2", width: 30}]])

      # Check column widths are tracked
      assert builder.column_widths[0] == 20
      assert builder.column_widths[1] == 30

      {:ok, content} = Builder.write_binary(builder)
      sheet_xml = xlsx_part(content, "xl/worksheets/sheet1.xml")
      assert sheet_xml =~ ~s(<col min="1" max="1" width="20.7109375")
      assert sheet_xml =~ ~s(<col min="2" max="2" width="30.7109375")
    end
  end

  defp shared_string_index(shared_strings, text) do
    ~r{<si>(.*?)</si>}
    |> Regex.scan(shared_strings, capture: :all_but_first)
    |> Enum.find_index(&(&1 == ["<t>#{text}</t>"]))
  end

  describe "format option conversions" do
    test "converts bold option correctly" do
      {:ok, content} =
        Builder.create()
        |> Builder.add_sheet("Test")
        |> Builder.add_rows([[{"Bold", format: [:bold]}]])
        |> Builder.write_binary()

      assert xlsx_part(content, "xl/styles.xml") =~ "<b/>"

      assert xlsx_part(content, "xl/worksheets/sheet1.xml") =~
               ~s(<c r="A1" s="1" t="s">)
    end

    test "converts multiple format options" do
      {:ok, content} =
        Builder.create()
        |> Builder.add_sheet("Test")
        |> Builder.add_rows([
          [
            {
              "Formatted",
              format: [
                :bold,
                :italic,
                {:font_size, 14},
                {:font_color, "#FF0000"},
                {:bg_color, "#FFFF00"}
              ]
            }
          ]
        ])
        |> Builder.write_binary()

      styles = xlsx_part(content, "xl/styles.xml")
      assert styles =~ ~s(<b/><i/><sz val="14"/><color rgb="FFFF0000"/>)
      assert styles =~ ~s(<fgColor rgb="FFFFFF00"/>)
    end

    test "handles alignment options" do
      {:ok, content} =
        Builder.create()
        |> Builder.add_sheet("Test")
        |> Builder.add_rows([
          [
            {"Left", format: [{:align, :left}]},
            {"Center", format: [{:align, :center}]},
            {"Right", format: [{:align, :right}]}
          ]
        ])
        |> Builder.write_binary()

      styles = xlsx_part(content, "xl/styles.xml")

      for alignment <- ["left", "center", "right"] do
        assert styles =~ ~s(<alignment horizontal="#{alignment}"/>)
      end
    end

    test "handles border options" do
      {:ok, content} =
        Builder.create()
        |> Builder.add_sheet("Test")
        |> Builder.add_rows([
          [
            {"Border All", format: [{:border, :thin}]},
            {"Border Top", format: [{:border_top, :medium}]}
          ]
        ])
        |> Builder.write_binary()

      styles = xlsx_part(content, "xl/styles.xml")
      assert styles =~ ~s(<left style="thin">)
      assert styles =~ ~s(<top style="medium">)
    end
  end
end
