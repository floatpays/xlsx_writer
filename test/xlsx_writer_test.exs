defmodule XlsxWriterTest do
  use ExUnit.Case
  doctest XlsxWriter

  describe "write/1" do
    test "write xlsx file" do
      bird_content = File.read!("bird.jpeg")

      sheets = [
        {"foobar",
         [
           {:write, 9, 0,
            {:string_with_format, "this is new", [{:align, :right}]}},
           {:write, 0, 0,
            {:string_with_format, "this is new", [:bold, {:align, :center}]}},
           {:write, 0, 1, {:float, 12.12}},
           {:write, 0, 3, {:image_path, "bird.jpeg"}},
           {:write, 1, 2, {:image, bird_content}},
           {:write, 2, 0, {:date, "2020-01-01"}},
           {:set_column_width, 0, 30},
           {:set_row_height, 0, 30}
         ]},
        {"zar", []}
      ]

      assert {:ok, content} = XlsxWriter.generate(sheets)

      assert <<80, _>> <> _ = content

      File.write!("test1.xlsx", content)
    end

    test "write simple file with some plain text data" do
      sheets = [
        {"sheet1",
         [
           {:write, 2, 1, {:string, ""}},
           {:write, 2, 0, {:string, "foo"}},
           {:write, 0, 1, {:string, "h2"}},
           {:write, 0, 0, {:string, "h1"}}
         ]}
      ]

      assert {:ok, content} = XlsxWriter.generate(sheets)

      File.write!("test2.xlsx", content)
    end

    test "write xlsx file with porcelain" do
      filename = "test2.xlsx"

      sheet1 =
        XlsxWriter.new_sheet("sheet number one")
        |> XlsxWriter.write(0, 0, "col1", format: [:bold])
        |> XlsxWriter.write(0, 1, "col2", format: [:bold, {:align, :center}])
        |> XlsxWriter.write(0, 2, "col3", format: [:bold, {:align, :right}])
        |> XlsxWriter.write(0, 3, nil)
        |> XlsxWriter.set_column_width(0, 40)
        |> XlsxWriter.set_column_width(3, 60)
        |> XlsxWriter.write(1, 0, "row 2 col 1")
        |> XlsxWriter.write(1, 1, 1.0)
        |> XlsxWriter.write_formula(1, 2, "=B2 + 2")
        |> XlsxWriter.write_formula(2, 1, "=PI()")
        |> XlsxWriter.write_image(3, 0, File.read!("bird.jpeg"))
        |> XlsxWriter.write(4, 3, 1)
        |> XlsxWriter.write(5, 3, DateTime.utc_now())
        |> XlsxWriter.write(6, 3, NaiveDateTime.utc_now())
        |> XlsxWriter.write(7, 3, Date.utc_today())
        |> XlsxWriter.write(8, 3, Decimal.new("20.12"))

      sheet2 =
        XlsxWriter.new_sheet("sheet number two")
        |> XlsxWriter.write(0, 0, "col1")

      {:ok, content} = XlsxWriter.generate([sheet1, sheet2])

      File.write!(filename, content)
    end

    test "write xlsx file with numeric format" do
      filename = "test2.xlsx"

      sheet1 =
        XlsxWriter.new_sheet("sheet number one")
        |> XlsxWriter.write(0, 0, 999.99, format: [
          {:num_format, "[$R] #,##0.00"}
        ])
        |> XlsxWriter.write(1, 0, 888, format: [{:num_format, "0,000.00"}])

      {:ok, content} = XlsxWriter.generate([sheet1])

      File.write!(filename, content)
    end

    test "write xlsx file with unsupported format" do
      assert_raise XlsxWriter.Error, fn ->
        XlsxWriter.new_sheet("sheet number one")
        |> XlsxWriter.write(0, 0, self())
      end
    end
  end

  describe "write_boolean/5" do
    test "generates valid xlsx with boolean values" do
      sheet =
        XlsxWriter.new_sheet("Boolean Test")
        |> XlsxWriter.write(0, 0, "Boolean Column", format: [:bold])
        |> XlsxWriter.write_boolean(1, 0, true)
        |> XlsxWriter.write_boolean(2, 0, false)
        |> XlsxWriter.write_boolean(3, 0, true, format: [:bold, {:align, :center}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "write_url/5" do
    test "generates valid xlsx with URLs" do
      sheet =
        XlsxWriter.new_sheet("URL Test")
        |> XlsxWriter.write(0, 0, "Links", format: [:bold])
        |> XlsxWriter.write_url(1, 0, "https://elixir-lang.org")
        |> XlsxWriter.write_url(2, 0, "https://hexdocs.pm", text: "Hex Docs")
        |> XlsxWriter.write_url(3, 0, "https://github.com", format: [:bold])
        |> XlsxWriter.write_url(4, 0, "https://anthropic.com",
          text: "Anthropic",
          format: [{:align, :center}]
        )

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "write_blank/4" do
    test "generates valid xlsx with blank cells" do
      sheet =
        XlsxWriter.new_sheet("Blank Test")
        |> XlsxWriter.write(0, 0, "Header 1", format: [:bold])
        |> XlsxWriter.write(0, 1, "Header 2", format: [:bold])
        |> XlsxWriter.write_blank(1, 0, format: [{:align, :center}])
        |> XlsxWriter.write_blank(1, 1, format: [:bold, {:align, :right}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "new data types integration" do
    test "generates xlsx with all new data types combined" do
      sheet =
        XlsxWriter.new_sheet("All Features")
        |> XlsxWriter.write(0, 0, "Type", format: [:bold])
        |> XlsxWriter.write(0, 1, "Value", format: [:bold])
        |> XlsxWriter.write(1, 0, "Boolean")
        |> XlsxWriter.write_boolean(1, 1, true)
        |> XlsxWriter.write(2, 0, "URL")
        |> XlsxWriter.write_url(2, 1, "https://example.com", text: "Example")
        |> XlsxWriter.write(3, 0, "Blank")
        |> XlsxWriter.write_blank(3, 1, format: [{:align, :center}])
        |> XlsxWriter.set_column_width(0, 20)
        |> XlsxWriter.set_column_width(1, 30)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
      assert byte_size(content) > 0
    end
  end

  describe "freeze_panes/3" do
    test "generates valid xlsx with frozen panes" do
      sheet =
        XlsxWriter.new_sheet("Frozen Panes")
        |> XlsxWriter.write(0, 0, "Header 1", format: [:bold])
        |> XlsxWriter.write(0, 1, "Header 2", format: [:bold])
        |> XlsxWriter.freeze_panes(1, 0)
        |> XlsxWriter.write(1, 0, "Data 1")
        |> XlsxWriter.write(1, 1, "Data 2")

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "hide_row/2 and hide_column/2" do
    test "generates valid xlsx with hidden row and column" do
      sheet =
        XlsxWriter.new_sheet("Hidden")
        |> XlsxWriter.write(0, 0, "Visible")
        |> XlsxWriter.write(1, 0, "Hidden Row")
        |> XlsxWriter.hide_row(1)
        |> XlsxWriter.write(0, 1, "Visible Col")
        |> XlsxWriter.write(0, 2, "Hidden Col")
        |> XlsxWriter.hide_column(2)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "set_autofilter/5" do
    test "generates valid xlsx with autofilter" do
      sheet =
        XlsxWriter.new_sheet("Autofilter")
        |> XlsxWriter.write(0, 0, "Name", format: [:bold])
        |> XlsxWriter.write(0, 1, "Age", format: [:bold])
        |> XlsxWriter.write(0, 2, "City", format: [:bold])
        |> XlsxWriter.set_autofilter(0, 0, 0, 2)
        |> XlsxWriter.write(1, 0, "Alice")
        |> XlsxWriter.write(1, 1, 30)
        |> XlsxWriter.write(1, 2, "NYC")

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "merge_range/7" do
    test "generates valid xlsx with merged cells" do
      sheet =
        XlsxWriter.new_sheet("Merged")
        |> XlsxWriter.merge_range(0, 0, 0, 3, "Title", format: [:bold, {:align, :center}])
        |> XlsxWriter.write(1, 0, "Col 1")
        |> XlsxWriter.write(1, 1, "Col 2")
        |> XlsxWriter.merge_range(2, 0, 4, 0, 100)
        |> XlsxWriter.merge_range(2, 1, 4, 1, true, format: [:bold])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "background colors" do
    test "generates valid xlsx with background colors" do
      sheet =
        XlsxWriter.new_sheet("Colors")
        |> XlsxWriter.write(0, 0, "Red", format: [{:bg_color, "#FF0000"}])
        |> XlsxWriter.write(0, 1, "Green", format: [{:bg_color, "#00FF00"}])
        |> XlsxWriter.write(0, 2, "Blue", format: [{:bg_color, "#0000FF"}])
        |> XlsxWriter.write(1, 0, "Yellow", format: [{:bg_color, "#FFFF00"}])
        |> XlsxWriter.write(1, 1, "Cyan", format: [{:bg_color, "#00FFFF"}])
        |> XlsxWriter.write(1, 2, "Magenta", format: [{:bg_color, "#FF00FF"}])
        |> XlsxWriter.write(2, 0, "Bold + Color", format: [:bold, {:bg_color, "#FFA500"}])
        |> XlsxWriter.write(2, 1, 100, format: [{:bg_color, "#90EE90"}, {:num_format, "$#,##0.00"}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "phase 1 features integration" do
    test "generates xlsx with all phase 1 features combined" do
      sheet =
        XlsxWriter.new_sheet("Phase 1 Features")
        # Merged header
        |> XlsxWriter.merge_range(0, 0, 0, 4, "Sales Report", format: [:bold, {:align, :center}])
        # Column headers with autofilter
        |> XlsxWriter.write(1, 0, "Product", format: [:bold])
        |> XlsxWriter.write(1, 1, "Q1", format: [:bold])
        |> XlsxWriter.write(1, 2, "Q2", format: [:bold])
        |> XlsxWriter.write(1, 3, "Q3", format: [:bold])
        |> XlsxWriter.write(1, 4, "Q4", format: [:bold])
        |> XlsxWriter.set_autofilter(1, 0, 1, 4)
        # Freeze header rows
        |> XlsxWriter.freeze_panes(2, 0)
        # Data
        |> XlsxWriter.write(2, 0, "Widget A")
        |> XlsxWriter.write(2, 1, 100)
        |> XlsxWriter.write(2, 2, 150)
        |> XlsxWriter.write(2, 3, 125)
        |> XlsxWriter.write(2, 4, 175)
        # Hidden row
        |> XlsxWriter.write(3, 0, "Hidden Product")
        |> XlsxWriter.hide_row(3)
        # Hidden column data
        |> XlsxWriter.write(2, 5, "Hidden Data")
        |> XlsxWriter.hide_column(5)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
      assert byte_size(content) > 0
    end
  end
end
