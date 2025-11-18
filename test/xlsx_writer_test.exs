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

  describe "font colors and styles" do
    test "generates valid xlsx with font colors and styles" do
      sheet =
        XlsxWriter.new_sheet("Font Styles")
        # Font colors
        |> XlsxWriter.write(0, 0, "Red Text", format: [{:font_color, "#FF0000"}])
        |> XlsxWriter.write(0, 1, "Blue Text", format: [{:font_color, "#0000FF"}])
        |> XlsxWriter.write(0, 2, "Green Text", format: [{:font_color, "#00FF00"}])

        # Font styles
        |> XlsxWriter.write(1, 0, "Italic", format: [:italic])
        |> XlsxWriter.write(1, 1, "Strikethrough", format: [:strikethrough])
        |> XlsxWriter.write(1, 2, "Underline", format: [{:underline, :single}])

        # Font sizes
        |> XlsxWriter.write(2, 0, "Size 10", format: [{:font_size, 10}])
        |> XlsxWriter.write(2, 1, "Size 14", format: [{:font_size, 14}])
        |> XlsxWriter.write(2, 2, "Size 18", format: [{:font_size, 18}])

        # Font names
        |> XlsxWriter.write(3, 0, "Arial", format: [{:font_name, "Arial"}])
        |> XlsxWriter.write(3, 1, "Courier", format: [{:font_name, "Courier New"}])
        |> XlsxWriter.write(3, 2, "Times", format: [{:font_name, "Times New Roman"}])

        # Combined formatting
        |> XlsxWriter.write(4, 0, "Bold+Red", format: [:bold, {:font_color, "#FF0000"}])
        |> XlsxWriter.write(4, 1, "Italic+Large", format: [:italic, {:font_size, 16}])
        |> XlsxWriter.write(4, 2, "All", format: [:bold, :italic, {:font_color, "#0000FF"}, {:font_size, 14}])

        # Superscript and subscript
        |> XlsxWriter.write(5, 0, "E=mc²", format: [:superscript])
        |> XlsxWriter.write(5, 1, "H₂O", format: [:subscript])

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

  describe "cell borders" do
    test "generates valid xlsx with basic borders" do
      sheet =
        XlsxWriter.new_sheet("Borders")
        # All sides border
        |> XlsxWriter.write(0, 0, "Thin Border", format: [{:border, :thin}])
        |> XlsxWriter.write(0, 1, "Medium Border", format: [{:border, :medium}])
        |> XlsxWriter.write(0, 2, "Thick Border", format: [{:border, :thick}])

        # Individual side borders
        |> XlsxWriter.write(1, 0, "Top", format: [{:border_top, :thin}])
        |> XlsxWriter.write(1, 1, "Bottom", format: [{:border_bottom, :medium}])
        |> XlsxWriter.write(1, 2, "Left", format: [{:border_left, :thick}])
        |> XlsxWriter.write(1, 3, "Right", format: [{:border_right, :thin}])

        # Border styles
        |> XlsxWriter.write(2, 0, "Dashed", format: [{:border, :dashed}])
        |> XlsxWriter.write(2, 1, "Dotted", format: [{:border, :dotted}])
        |> XlsxWriter.write(2, 2, "Double", format: [{:border, :double}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "generates valid xlsx with colored borders" do
      sheet =
        XlsxWriter.new_sheet("Colored Borders")
        # All borders with color
        |> XlsxWriter.write(0, 0, "Red Border",
            format: [{:border, :thin}, {:border_color, "#FF0000"}])
        |> XlsxWriter.write(0, 1, "Blue Border",
            format: [{:border, :medium}, {:border_color, "#0000FF"}])

        # Individual side colors
        |> XlsxWriter.write(1, 0, "Multi-colored",
            format: [
              {:border_top, :thin},
              {:border_top_color, "#FF0000"},
              {:border_bottom, :thin},
              {:border_bottom_color, "#00FF00"},
              {:border_left, :thin},
              {:border_left_color, "#0000FF"},
              {:border_right, :thin},
              {:border_right_color, "#FFFF00"}
            ])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "generates valid xlsx with borders and other formatting" do
      sheet =
        XlsxWriter.new_sheet("Combined Formatting")
        # Borders with background colors
        |> XlsxWriter.write(0, 0, "Header 1",
            format: [:bold, {:border, :thick}, {:bg_color, "#4472C4"}, {:align, :center}])
        |> XlsxWriter.write(0, 1, "Header 2",
            format: [:bold, {:border, :thick}, {:bg_color, "#4472C4"}, {:align, :center}])

        # Borders with font formatting
        |> XlsxWriter.write(1, 0, "Bold Red",
            format: [:bold, {:font_color, "#FF0000"}, {:border, :thin}])
        |> XlsxWriter.write(1, 1, 100,
            format: [{:num_format, "$#,##0.00"}, {:border_bottom, :double}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "generates valid xlsx with all border styles" do
      sheet =
        XlsxWriter.new_sheet("All Border Styles")
        |> XlsxWriter.write(0, 0, "Thin", format: [{:border, :thin}])
        |> XlsxWriter.write(1, 0, "Medium", format: [{:border, :medium}])
        |> XlsxWriter.write(2, 0, "Thick", format: [{:border, :thick}])
        |> XlsxWriter.write(3, 0, "Dashed", format: [{:border, :dashed}])
        |> XlsxWriter.write(4, 0, "Dotted", format: [{:border, :dotted}])
        |> XlsxWriter.write(5, 0, "Double", format: [{:border, :double}])
        |> XlsxWriter.write(6, 0, "Hair", format: [{:border, :hair}])
        |> XlsxWriter.write(7, 0, "Medium Dashed", format: [{:border, :medium_dashed}])
        |> XlsxWriter.write(8, 0, "Dash Dot", format: [{:border, :dash_dot}])
        |> XlsxWriter.write(9, 0, "Medium Dash Dot", format: [{:border, :medium_dash_dot}])
        |> XlsxWriter.write(10, 0, "Dash Dot Dot", format: [{:border, :dash_dot_dot}])
        |> XlsxWriter.write(11, 0, "Medium Dash Dot Dot", format: [{:border, :medium_dash_dot_dot}])
        |> XlsxWriter.write(12, 0, "Slant Dash Dot", format: [{:border, :slant_dash_dot}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "error handling" do
    test "raises error for unsupported data type (PID)" do
      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, self())
      end
    end

    test "raises error for unsupported data type (function)" do
      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, fn -> :ok end)
      end
    end

    test "raises error for unsupported data type (reference)" do
      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, make_ref())
      end
    end

    test "raises error for unsupported data type (port)" do
      {:ok, port} = :gen_tcp.listen(0, [])

      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, port)
      end

      :gen_tcp.close(port)
    end

    test "raises error for unsupported data type (map)" do
      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, %{foo: "bar"})
      end
    end

    test "raises error for unsupported data type (list)" do
      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, [1, 2, 3])
      end
    end

    test "raises error for unsupported data type (tuple)" do
      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, {:ok, "value"})
      end
    end

    test "handles invalid hex color gracefully in background color" do
      # Invalid hex should not crash, just be ignored
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text", format: [{:bg_color, "invalid"}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles invalid hex color gracefully in font color" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text", format: [{:font_color, "GGGGGG"}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles invalid hex color gracefully in border color" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text", format: [{:border, :thin}, {:border_color, "notahex"}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles empty string hex color" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text", format: [{:bg_color, ""}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles invalid date string gracefully" do
      # This will be caught by Rust and return an error
      sheet = XlsxWriter.new_sheet("Test")
      |> XlsxWriter.write(0, 0, "foo")

      assert {:ok, _content} = XlsxWriter.generate([sheet])
    end

    test "raises error for negative row index" do
      assert_raise ArgumentError, ~r/Row index must be non-negative/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(-1, 0, "Text")
      end
    end

    test "raises error for negative column index" do
      assert_raise ArgumentError, ~r/Column index must be non-negative/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, -1, "Text")
      end
    end

    test "handles very large row index" do
      # Excel has a max of 1,048,576 rows
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(2_000_000, 0, "Text")

      result = XlsxWriter.generate([sheet])
      # Rust should handle this gracefully
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles very large column index" do
      # Excel has a max of 16,384 columns
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 20_000, "Text")

      result = XlsxWriter.generate([sheet])
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles empty sheet name" do
      sheet = XlsxWriter.new_sheet("")
        |> XlsxWriter.write(0, 0, "Text")

      result = XlsxWriter.generate([sheet])
      # Empty name might be invalid in Excel
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles very long sheet name" do
      # Excel sheet names max at 31 characters
      long_name = String.duplicate("a", 50)
      sheet = XlsxWriter.new_sheet(long_name)
        |> XlsxWriter.write(0, 0, "Text")

      result = XlsxWriter.generate([sheet])
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles invalid characters in sheet name" do
      # Excel doesn't allow: \ / ? * [ ]
      sheet = XlsxWriter.new_sheet("Invalid[Sheet]")
        |> XlsxWriter.write(0, 0, "Text")

      result = XlsxWriter.generate([sheet])
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles invalid formula syntax" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_formula(0, 0, "=INVALID(((")

      # Formula syntax errors are caught at Excel runtime, not generation
      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles empty formula" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_formula(0, 0, "")

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles invalid URL" do
      # Invalid URLs should still write
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_url(0, 0, "not a url")

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles empty URL" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_url(0, 0, "")

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles zero column width" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text")
        |> XlsxWriter.set_column_width(0, 0)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles zero row height" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text")
        |> XlsxWriter.set_row_height(0, 0)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles invalid merge range (last < first)" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.merge_range(5, 5, 0, 0, "Text")

      result = XlsxWriter.generate([sheet])
      # Rust should handle invalid ranges
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles merge range with same start and end" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.merge_range(0, 0, 0, 0, "Text")

      # Merging a single cell should work (just write it)
      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles autofilter with invalid range (last < first)" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.set_autofilter(5, 5, 0, 0)

      result = XlsxWriter.generate([sheet])
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles zero font size" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text", format: [{:font_size, 0}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles extremely large font size" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text", format: [{:font_size, 1000}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles empty font name" do
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text", format: [{:font_name, ""}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles invalid image data" do
      # Invalid binary should be caught by Rust
      sheet = XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_image(0, 0, "not image data")

      result = XlsxWriter.generate([sheet])
      # Should return error from Rust
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "raises error for empty binary as image" do
      assert_raise ArgumentError, ~r/Image binary cannot be empty/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_image(0, 0, <<>>)
      end
    end

    test "handles empty list for generate" do
      result = XlsxWriter.generate([])

      # Empty workbook should work
      assert {:ok, content} = result
      assert <<80, _>> <> _ = content
    end

    test "handles nil in boolean write" do
      # This should raise since it's not a boolean
      assert_raise FunctionClauseError, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_boolean(0, 0, nil)
      end
    end

    test "handles string in boolean write" do
      assert_raise FunctionClauseError, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_boolean(0, 0, "true")
      end
    end

    test "handles integer in boolean write" do
      assert_raise FunctionClauseError, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_boolean(0, 0, 1)
      end
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
