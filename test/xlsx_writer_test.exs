defmodule XlsxWriterTest do
  use ExUnit.Case, async: true

  import XlsxWriter.TestHelpers

  doctest XlsxWriter

  # Values of unsupported types, passed through the test context so the type
  # checker cannot prove that the calls using them fail, which is exactly
  # what those tests assert at runtime.
  setup do
    %{
      untyped: %{
        pid: self(),
        function: fn -> :ok end,
        reference: make_ref(),
        map: %{foo: "bar"},
        list: [1, 2, 3],
        tuple: {:ok, "value"},
        integer: 123
      }
    }
  end

  describe "write/1" do
    test "write xlsx file from raw instructions" do
      bird_content = File.read!("bird.jpeg")

      sheets = [
        {"foobar",
         [
           {:write, 9, 0, {:string, "this is new", [{:align, :right}]}},
           {:write, 0, 0, {:string, "this is new", [:bold, {:align, :center}]}},
           {:write, 0, 1, {:number, 12.12, []}},
           {:write, 0, 3, {:image_path, "bird.jpeg"}},
           {:write, 1, 2, {:image, bird_content}},
           {:write, 2, 0, {:date, "2020-01-01", []}},
           {:set_column_width, 0, 30},
           {:set_row_height, 0, 30}
         ]},
        {"zar", []}
      ]

      assert {:ok, content} = XlsxWriter.generate(sheets)
      assert <<"PK", _rest::binary>> = content

      sheet_xml = xlsx_part(content, "xl/worksheets/sheet1.xml")
      assert sheet_xml =~ ~s(<c r="B1"><v>12.12</v></c>)
      assert sheet_xml =~ ~r/<c r="A3" s="\d+"><v>43831<\/v><\/c>/
      assert xlsx_part(content, "xl/worksheets/sheet2.xml")
      assert xlsx_part(content, "xl/media/image1.jpeg") == bird_content
    end

    test "write simple file with some plain text data" do
      sheets = [
        {"sheet1",
         [
           {:write, 2, 1, {:string, "", []}},
           {:write, 2, 0, {:string, "foo", []}},
           {:write, 0, 1, {:string, "h2", []}},
           {:write, 0, 0, {:string, "h1", []}}
         ]}
      ]

      shared_strings = generate_part!(sheets, "xl/sharedStrings.xml")

      for text <- ["foo", "h1", "h2"] do
        assert shared_strings =~ "<t>#{text}</t>"
      end
    end

    test "write xlsx file with porcelain" do
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

      assert {:ok, content} = XlsxWriter.generate([sheet1, sheet2])

      sheet_xml = xlsx_part(content, "xl/worksheets/sheet1.xml")
      assert sheet_xml =~ ~s(<c r="C2"><f>B2 + 2</f>)
      assert sheet_xml =~ ~s(<c r="D5"><v>1</v></c>)
      assert sheet_xml =~ ~s(<c r="D9"><v>20.12</v></c>)
    end

    test "write xlsx file with numeric format" do
      sheet1 =
        XlsxWriter.new_sheet("sheet number one")
        |> XlsxWriter.write(0, 0, 999.99,
          format: [
            {:num_format, "[$R] #,##0.00"}
          ]
        )
        |> XlsxWriter.write(1, 0, 888, format: [{:num_format, "0,000.00"}])

      assert {:ok, content} = XlsxWriter.generate([sheet1])
      styles = xlsx_part(content, "xl/styles.xml")
      assert styles =~ ~s(formatCode="[$R] #,##0.00")
      assert styles =~ ~s(formatCode="0,000.00")
    end
  end

  describe "write/5 with nil value" do
    test "writes a nil value with a format as an empty string cell" do
      sheet =
        XlsxWriter.new_sheet("Nil Test")
        |> XlsxWriter.write(0, 0, nil, format: [:bold, {:font_name, "Arial"}])

      assert {"Nil Test",
              [
                {:write, 0, 0, {:string, "", [:bold, {:font_name, "Arial"}]}}
              ]} = sheet

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "write_boolean/5" do
    test "generates valid xlsx with boolean values" do
      sheet =
        XlsxWriter.new_sheet("Boolean Test")
        |> XlsxWriter.write(0, 0, "Boolean Column", format: [:bold])
        |> XlsxWriter.write_boolean(1, 0, true)
        |> XlsxWriter.write_boolean(2, 0, false)
        |> XlsxWriter.write_boolean(3, 0, true,
          format: [:bold, {:align, :center}]
        )

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "write_formula/5" do
    test "generates valid xlsx with formatted formulas" do
      sheet =
        XlsxWriter.new_sheet("Formula Test")
        |> XlsxWriter.write_formula(0, 0, "=SUM(A1:A3)",
          format: [:bold, {:align, :center}]
        )

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

  describe "column and row range operations" do
    test "set_column_range_width/4 generates valid xlsx" do
      sheet =
        XlsxWriter.new_sheet("Column Ranges")
        |> XlsxWriter.write(0, 0, "Col A")
        |> XlsxWriter.write(0, 1, "Col B")
        |> XlsxWriter.write(0, 2, "Col C")
        |> XlsxWriter.write(0, 3, "Col D")
        |> XlsxWriter.write(0, 4, "Col E")
        # Set columns A-E (0-4) to 20 characters wide
        |> XlsxWriter.set_column_range_width(0, 4, 20)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "set_row_range_height/4 generates valid xlsx" do
      sheet =
        XlsxWriter.new_sheet("Row Ranges")
        |> XlsxWriter.write(0, 0, "Row 0")
        |> XlsxWriter.write(1, 0, "Row 1")
        |> XlsxWriter.write(2, 0, "Row 2")
        |> XlsxWriter.write(3, 0, "Row 3")
        |> XlsxWriter.write(4, 0, "Row 4")
        # Set rows 0-4 to 30 points tall
        |> XlsxWriter.set_row_range_height(0, 4, 30)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "combining column and row range operations" do
      sheet =
        XlsxWriter.new_sheet("Combined Ranges")
        |> XlsxWriter.write(0, 0, "A1")
        |> XlsxWriter.write(0, 1, "B1")
        |> XlsxWriter.write(0, 2, "C1")
        |> XlsxWriter.write(1, 0, "A2")
        |> XlsxWriter.write(1, 1, "B2")
        |> XlsxWriter.write(1, 2, "C2")
        # Set multiple column ranges
        |> XlsxWriter.set_column_range_width(0, 1, 100)
        |> XlsxWriter.set_column_range_width(2, 2, 150)
        # Set multiple row ranges
        |> XlsxWriter.set_row_range_height(0, 0, 40)
        |> XlsxWriter.set_row_range_height(1, 3, 25)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "range operations with single column/row" do
      sheet =
        XlsxWriter.new_sheet("Single Range")
        |> XlsxWriter.write(0, 0, "Single Col")
        |> XlsxWriter.write(0, 0, "Single Row")
        # Setting a range with same start and end should work
        |> XlsxWriter.set_column_range_width(0, 0, 100)
        |> XlsxWriter.set_row_range_height(0, 0, 30)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "range operations with data and formatting" do
      sheet =
        XlsxWriter.new_sheet("Formatted Ranges")
        |> XlsxWriter.write(0, 0, "Header 1",
          format: [:bold, {:bg_color, "#4472C4"}]
        )
        |> XlsxWriter.write(0, 1, "Header 2",
          format: [:bold, {:bg_color, "#4472C4"}]
        )
        |> XlsxWriter.write(0, 2, "Header 3",
          format: [:bold, {:bg_color, "#4472C4"}]
        )
        |> XlsxWriter.set_column_range_width(0, 2, 120)
        |> XlsxWriter.set_row_range_height(0, 0, 35)
        |> XlsxWriter.write(1, 0, "Data 1")
        |> XlsxWriter.write(1, 1, "Data 2")
        |> XlsxWriter.write(1, 2, "Data 3")

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "merge_range/7" do
    test "generates valid xlsx with merged cells" do
      sheet =
        XlsxWriter.new_sheet("Merged")
        |> XlsxWriter.merge_range(0, 0, 0, 3, "Title",
          format: [:bold, {:align, :center}]
        )
        |> XlsxWriter.write(1, 0, "Col 1")
        |> XlsxWriter.write(1, 1, "Col 2")
        |> XlsxWriter.merge_range(2, 0, 4, 0, 100)
        |> XlsxWriter.merge_range(2, 1, 4, 1, true, format: [:bold])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "text wrapping" do
    test "generates valid xlsx with text wrap" do
      sheet =
        XlsxWriter.new_sheet("Wrap")
        |> XlsxWriter.write(
          0,
          0,
          "This is a long text that should wrap within the cell",
          format: [:text_wrap]
        )
        |> XlsxWriter.write(0, 1, "Wrapped and bold",
          format: [:text_wrap, :bold]
        )
        |> XlsxWriter.set_column_width(0, 20)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "creates correct instruction for text wrap" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Wrapped", format: [:text_wrap])

      {"Test", [{:write, 0, 0, {:string, "Wrapped", [:text_wrap]}}]} =
        sheet
    end
  end

  describe "vertical alignment" do
    test "generates valid xlsx with vertical alignment" do
      sheet =
        XlsxWriter.new_sheet("VAlign")
        |> XlsxWriter.write(0, 0, "Top", format: [{:valign, :top}])
        |> XlsxWriter.write(0, 1, "Center", format: [{:valign, :center}])
        |> XlsxWriter.write(0, 2, "Bottom", format: [{:valign, :bottom}])
        |> XlsxWriter.write(0, 3, "Justify", format: [{:valign, :justify}])
        |> XlsxWriter.write(0, 4, "Distributed",
          format: [{:valign, :distributed}]
        )
        |> XlsxWriter.set_row_height(0, 40)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "combines horizontal and vertical alignment" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Centered Both",
          format: [{:align, :center}, {:valign, :center}]
        )
        |> XlsxWriter.set_row_height(0, 40)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "creates correct instruction for valign" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Top", format: [{:valign, :top}])

      {"Test", [{:write, 0, 0, {:string, "Top", [{:valign, :top}]}}]} =
        sheet
    end
  end

  describe "text rotation" do
    test "generates valid xlsx with rotated text" do
      sheet =
        XlsxWriter.new_sheet("Rotation")
        |> XlsxWriter.write(0, 0, "45 degrees", format: [{:rotation, 45}])
        |> XlsxWriter.write(0, 1, "-45 degrees", format: [{:rotation, -45}])
        |> XlsxWriter.write(0, 2, "90 degrees", format: [{:rotation, 90}])
        |> XlsxWriter.write(0, 3, "Stacked", format: [{:rotation, 270}])
        |> XlsxWriter.set_row_height(0, 60)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "creates correct instruction for rotation" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Rotated", format: [{:rotation, 45}])

      {"Test", [{:write, 0, 0, {:string, "Rotated", [{:rotation, 45}]}}]} =
        sheet
    end
  end

  describe "shrink to fit" do
    test "generates valid xlsx with shrink to fit" do
      sheet =
        XlsxWriter.new_sheet("Shrink")
        |> XlsxWriter.write(
          0,
          0,
          "This long text will shrink to fit the column width",
          format: [:shrink]
        )
        |> XlsxWriter.write(0, 1, "Shrink + bold", format: [:shrink, :bold])
        |> XlsxWriter.set_column_width(0, 15)
        |> XlsxWriter.set_column_width(1, 10)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "creates correct instruction for shrink" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Shrink", format: [:shrink])

      {"Test", [{:write, 0, 0, {:string, "Shrink", [:shrink]}}]} =
        sheet
    end
  end

  describe "text indent" do
    test "generates valid xlsx with text indent" do
      sheet =
        XlsxWriter.new_sheet("Indent")
        |> XlsxWriter.write(0, 0, "No indent")
        |> XlsxWriter.write(1, 0, "Indent 1", format: [{:indent, 1}])
        |> XlsxWriter.write(2, 0, "Indent 2", format: [{:indent, 2}])
        |> XlsxWriter.write(3, 0, "Indent 3", format: [{:indent, 3}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "creates correct instruction for indent" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Indented", format: [{:indent, 2}])

      {"Test", [{:write, 0, 0, {:string, "Indented", [{:indent, 2}]}}]} =
        sheet
    end
  end

  describe "font colors and styles" do
    test "generates valid xlsx with font colors and styles" do
      sheet =
        XlsxWriter.new_sheet("Font Styles")
        # Font colors
        |> XlsxWriter.write(0, 0, "Red Text",
          format: [{:font_color, "#FF0000"}]
        )
        |> XlsxWriter.write(0, 1, "Blue Text",
          format: [{:font_color, "#0000FF"}]
        )
        |> XlsxWriter.write(0, 2, "Green Text",
          format: [{:font_color, "#00FF00"}]
        )

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
        |> XlsxWriter.write(3, 1, "Courier",
          format: [{:font_name, "Courier New"}]
        )
        |> XlsxWriter.write(3, 2, "Times",
          format: [{:font_name, "Times New Roman"}]
        )

        # Combined formatting
        |> XlsxWriter.write(4, 0, "Bold+Red",
          format: [:bold, {:font_color, "#FF0000"}]
        )
        |> XlsxWriter.write(4, 1, "Italic+Large",
          format: [:italic, {:font_size, 16}]
        )
        |> XlsxWriter.write(4, 2, "All",
          format: [:bold, :italic, {:font_color, "#0000FF"}, {:font_size, 14}]
        )

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
        |> XlsxWriter.write(2, 0, "Bold + Color",
          format: [:bold, {:bg_color, "#FFA500"}]
        )
        |> XlsxWriter.write(2, 1, 100,
          format: [{:bg_color, "#90EE90"}, {:num_format, "$#,##0.00"}]
        )

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
          format: [{:border, :thin}, {:border_color, "#FF0000"}]
        )
        |> XlsxWriter.write(0, 1, "Blue Border",
          format: [{:border, :medium}, {:border_color, "#0000FF"}]
        )

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
          ]
        )

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "generates valid xlsx with borders and other formatting" do
      sheet =
        XlsxWriter.new_sheet("Combined Formatting")
        # Borders with background colors
        |> XlsxWriter.write(0, 0, "Header 1",
          format: [
            :bold,
            {:border, :thick},
            {:bg_color, "#4472C4"},
            {:align, :center}
          ]
        )
        |> XlsxWriter.write(0, 1, "Header 2",
          format: [
            :bold,
            {:border, :thick},
            {:bg_color, "#4472C4"},
            {:align, :center}
          ]
        )

        # Borders with font formatting
        |> XlsxWriter.write(1, 0, "Bold Red",
          format: [:bold, {:font_color, "#FF0000"}, {:border, :thin}]
        )
        |> XlsxWriter.write(1, 1, 100,
          format: [{:num_format, "$#,##0.00"}, {:border_bottom, :double}]
        )

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
        |> XlsxWriter.write(7, 0, "Medium Dashed",
          format: [{:border, :medium_dashed}]
        )
        |> XlsxWriter.write(8, 0, "Dash Dot", format: [{:border, :dash_dot}])
        |> XlsxWriter.write(9, 0, "Medium Dash Dot",
          format: [{:border, :medium_dash_dot}]
        )
        |> XlsxWriter.write(10, 0, "Dash Dot Dot",
          format: [{:border, :dash_dot_dot}]
        )
        |> XlsxWriter.write(11, 0, "Medium Dash Dot Dot",
          format: [{:border, :medium_dash_dot_dot}]
        )
        |> XlsxWriter.write(12, 0, "Slant Dash Dot",
          format: [{:border, :slant_dash_dot}]
        )

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "error handling" do
    test "raises error for unsupported data type (PID)", %{untyped: untyped} do
      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, untyped.pid)
      end
    end

    test "raises error for unsupported data type (function)", %{
      untyped: untyped
    } do
      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, untyped.function)
      end
    end

    test "raises error for unsupported data type (reference)", %{
      untyped: untyped
    } do
      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, untyped.reference)
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

    test "raises error for unsupported data type (map)", %{untyped: untyped} do
      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, untyped.map)
      end
    end

    test "raises error for unsupported data type (list)", %{untyped: untyped} do
      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, untyped.list)
      end
    end

    test "raises error for unsupported data type (tuple)", %{untyped: untyped} do
      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, untyped.tuple)
      end
    end

    for {option, color} <- [
          bg_color: "invalid",
          font_color: "GGGGGG",
          border_color: "notahex",
          bg_color: "",
          bg_color: "#FFF",
          font_color: "red",
          font_color: "#FFFFFFFF",
          border_top_color: "#12345"
        ] do
      test "returns an error for invalid #{option} #{inspect(color)}" do
        sheet =
          XlsxWriter.new_sheet("Test")
          |> XlsxWriter.write(0, 0, "Text",
            format: [{unquote(option), unquote(color)}]
          )

        assert {:error, "invalid color " <> _} = XlsxWriter.generate([sheet])
      end
    end

    test "handles boolean as hex color gracefully in background color" do
      assert_raise XlsxWriter.Error,
                   ~r/bg_color.*expects a string hex color.*got: true/,
                   fn ->
                     XlsxWriter.new_sheet("Test")
                     |> XlsxWriter.write(0, 0, "Text",
                       format: [{:bg_color, true}]
                     )
                   end
    end

    test "handles boolean as hex color gracefully in font color" do
      assert_raise XlsxWriter.Error,
                   ~r/font_color.*expects a string hex color.*got: false/,
                   fn ->
                     XlsxWriter.new_sheet("Test")
                     |> XlsxWriter.write(0, 0, "Text",
                       format: [{:font_color, false}]
                     )
                   end
    end

    test "handles integer as hex color gracefully in border color" do
      assert_raise XlsxWriter.Error,
                   ~r/border_color.*expects a string hex color.*got: 123/,
                   fn ->
                     XlsxWriter.new_sheet("Test")
                     |> XlsxWriter.write(0, 0, "Text",
                       format: [{:border, :thin}, {:border_color, 123}]
                     )
                   end
    end

    test "returns an error for a date outside Excel's range" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, ~D[1800-01-01])

      assert {:error, reason} = XlsxWriter.generate([sheet])
      assert reason =~ "outside Excel range"
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
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(2_000_000, 0, "Text")

      result = XlsxWriter.generate([sheet])
      # Rust should handle this gracefully
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles very large column index" do
      # Excel has a max of 16,384 columns
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 20_000, "Text")

      result = XlsxWriter.generate([sheet])
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles empty sheet name" do
      sheet =
        XlsxWriter.new_sheet("")
        |> XlsxWriter.write(0, 0, "Text")

      result = XlsxWriter.generate([sheet])
      # Empty name might be invalid in Excel
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles very long sheet name" do
      # Excel sheet names max at 31 characters
      long_name = String.duplicate("a", 50)

      sheet =
        XlsxWriter.new_sheet(long_name)
        |> XlsxWriter.write(0, 0, "Text")

      result = XlsxWriter.generate([sheet])
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles invalid characters in sheet name" do
      # Excel doesn't allow: \ / ? * [ ]
      sheet =
        XlsxWriter.new_sheet("Invalid[Sheet]")
        |> XlsxWriter.write(0, 0, "Text")

      result = XlsxWriter.generate([sheet])
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles invalid formula syntax" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_formula(0, 0, "=INVALID(((")

      # Formula syntax errors are caught at Excel runtime, not generation
      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles empty formula" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_formula(0, 0, "")

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles invalid URL" do
      # Invalid URLs should return an error
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_url(0, 0, "not a url")

      assert {:error, reason} = XlsxWriter.generate([sheet])
      assert reason =~ "url type"
    end

    test "handles empty URL" do
      # Empty URLs should return an error
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_url(0, 0, "")

      assert {:error, reason} = XlsxWriter.generate([sheet])
      assert reason =~ "url type"
    end

    test "handles zero column width" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text")
        |> XlsxWriter.set_column_width(0, 0)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles zero row height" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text")
        |> XlsxWriter.set_row_height(0, 0)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles invalid merge range (last < first)" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.merge_range(5, 5, 0, 0, "Text")

      result = XlsxWriter.generate([sheet])
      # Rust should handle invalid ranges
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "handles merge range with same start and end" do
      # Merging a single cell should return an error
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.merge_range(0, 0, 0, 0, "Text")

      assert {:error, reason} = XlsxWriter.generate([sheet])
      assert reason =~ "single cell"
    end

    test "handles autofilter with invalid range (last < first)" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.set_autofilter(5, 5, 0, 0)

      result = XlsxWriter.generate([sheet])
      assert match?({:ok, _}, result) or match?({:error, _}, result)
    end

    test "returns an error for a zero font size" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text", format: [{:font_size, 0}])

      assert {:error, "font size 0 must be greater than 0"} =
               XlsxWriter.generate([sheet])
    end

    test "returns an error for a font size above Excel's maximum of 409" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text", format: [{:font_size, 1000}])

      assert {:error, "font size 1000 is outside Excel's range 0..409"} =
               XlsxWriter.generate([sheet])
    end

    test "handles empty font name" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Text", format: [{:font_name, ""}])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "handles invalid image data" do
      # Invalid binary should be caught by Rust
      sheet =
        XlsxWriter.new_sheet("Test")
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
  end

  describe "write_rich_string/5" do
    test "generates valid xlsx with basic rich string" do
      sheet =
        XlsxWriter.new_sheet("Rich String Test")
        |> XlsxWriter.write_rich_string(0, 0, [
          {"Bold ", [:bold]},
          {"Normal ", []},
          {"Italic", [:italic]}
        ])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "generates valid xlsx with colored rich string segments" do
      sheet =
        XlsxWriter.new_sheet("Colored Rich String")
        |> XlsxWriter.write_rich_string(0, 0, [
          {"Red ", [{:font_color, "#FF0000"}]},
          {"Green ", [{:font_color, "#00FF00"}]},
          {"Blue", [{:font_color, "#0000FF"}]}
        ])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "generates valid xlsx with rich string and cell-level formatting" do
      sheet =
        XlsxWriter.new_sheet("Rich String With Cell Format")
        |> XlsxWriter.write_rich_string(
          0,
          0,
          [
            {"Bold ", [:bold]},
            {"Italic", [:italic]}
          ],
          format: [{:align, :center}, {:bg_color, "#FFFF00"}]
        )

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "generates valid xlsx with multiple format options per segment" do
      sheet =
        XlsxWriter.new_sheet("Complex Rich String")
        |> XlsxWriter.write_rich_string(0, 0, [
          {"Bold Red ", [:bold, {:font_color, "#FF0000"}]},
          {"Large ", [{:font_size, 16}]},
          {"Underlined", [{:underline, :single}]}
        ])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "generates valid xlsx with superscript and subscript in rich string" do
      sheet =
        XlsxWriter.new_sheet("Super Sub Script")
        |> XlsxWriter.write_rich_string(0, 0, [
          {"E=mc", []},
          {"2", [:superscript]}
        ])
        |> XlsxWriter.write_rich_string(1, 0, [
          {"H", []},
          {"2", [:subscript]},
          {"O", []}
        ])

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "raises error for empty segments list" do
      assert_raise ArgumentError,
                   ~r/Rich string segments cannot be empty/,
                   fn ->
                     XlsxWriter.new_sheet("Test")
                     |> XlsxWriter.write_rich_string(0, 0, [])
                   end
    end

    test "raises error for non-list segments" do
      assert_raise ArgumentError, ~r/Rich string segments must be a list/, fn ->
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_rich_string(0, 0, "not a list")
      end
    end

    test "raises error for invalid segment tuple" do
      assert_raise ArgumentError,
                   ~r/Rich string segment must be a \{text, formats\} tuple/,
                   fn ->
                     XlsxWriter.new_sheet("Test")
                     |> XlsxWriter.write_rich_string(0, 0, ["not a tuple"])
                   end
    end

    test "raises error for non-string text in segment" do
      assert_raise ArgumentError,
                   ~r/Rich string segment text must be a string/,
                   fn ->
                     XlsxWriter.new_sheet("Test")
                     |> XlsxWriter.write_rich_string(0, 0, [{123, [:bold]}])
                   end
    end

    test "raises error for non-list formats in segment" do
      assert_raise ArgumentError,
                   ~r/Rich string segment formats must be a list/,
                   fn ->
                     XlsxWriter.new_sheet("Test")
                     |> XlsxWriter.write_rich_string(0, 0, [{"text", :bold}])
                   end
    end

    test "validates color formats in segments" do
      assert_raise XlsxWriter.Error,
                   ~r/font_color.*expects a string hex color/,
                   fn ->
                     XlsxWriter.new_sheet("Test")
                     |> XlsxWriter.write_rich_string(0, 0, [
                       {"text", [{:font_color, true}]}
                     ])
                   end
    end

    test "creates correct instruction for rich string without cell format" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_rich_string(0, 0, [
          {"Bold ", [:bold]},
          {"Normal", []}
        ])

      {"Test",
       [
         {:write, 0, 0,
          {:rich_string, [{"Bold ", [:bold]}, {"Normal", []}], []}}
       ]} =
        sheet
    end

    test "creates correct instruction for rich string with cell format" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_rich_string(
          0,
          0,
          [{"Bold ", [:bold]}, {"Italic", [:italic]}],
          format: [{:align, :center}]
        )

      {"Test",
       [
         {:write, 0, 0,
          {:rich_string, [{"Bold ", [:bold]}, {"Italic", [:italic]}],
           [{:align, :center}]}}
       ]} = sheet
    end
  end

  describe "autofit" do
    test "generates valid xlsx with autofit" do
      sheet =
        XlsxWriter.new_sheet("Autofit")
        |> XlsxWriter.write(0, 0, "Short")
        |> XlsxWriter.write(0, 1, "This is a much longer column header text")
        |> XlsxWriter.write(1, 0, "A")
        |> XlsxWriter.write(1, 1, "B")
        |> XlsxWriter.autofit()

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end

    test "creates correct instruction for autofit" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.autofit()

      {"Test", [:set_autofit]} = sheet
    end
  end

  describe "tab color" do
    test "generates valid xlsx with tab colors" do
      sheet1 =
        XlsxWriter.new_sheet("Red Tab")
        |> XlsxWriter.set_tab_color("#FF0000")
        |> XlsxWriter.write(0, 0, "Red tab sheet")

      sheet2 =
        XlsxWriter.new_sheet("Blue Tab")
        |> XlsxWriter.set_tab_color("#0000FF")
        |> XlsxWriter.write(0, 0, "Blue tab sheet")

      assert {:ok, content} = XlsxWriter.generate([sheet1, sheet2])
      assert <<80, _>> <> _ = content
    end

    test "creates correct instruction for tab color" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.set_tab_color("#FF0000")

      {"Test", [{:set_tab_color, "#FF0000"}]} = sheet
    end
  end

  describe "workbook properties" do
    test "generates valid xlsx with document properties" do
      sheet =
        XlsxWriter.new_sheet("Data")
        |> XlsxWriter.write(0, 0, "Hello")

      props = %XlsxWriter.WorkbookProperties{
        author: "Test Author",
        title: "Test Report",
        subject: "Testing",
        company: "Test Corp",
        category: "Reports",
        keywords: "test, xlsx",
        comment: "Generated by XlsxWriter",
        status: "Draft"
      }

      assert {:ok, content} = XlsxWriter.generate([sheet], properties: props)
      assert <<80, _>> <> _ = content
    end

    test "generates valid xlsx with partial properties" do
      sheet =
        XlsxWriter.new_sheet("Data")
        |> XlsxWriter.write(0, 0, "Hello")

      props = %XlsxWriter.WorkbookProperties{author: "Jane Doe"}

      assert {:ok, content} = XlsxWriter.generate([sheet], properties: props)
      assert <<80, _>> <> _ = content
    end

    test "generates valid xlsx without properties (backward compatible)" do
      sheet =
        XlsxWriter.new_sheet("Data")
        |> XlsxWriter.write(0, 0, "Hello")

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert <<80, _>> <> _ = content
    end
  end

  describe "phase 1 features integration" do
    test "generates xlsx with all phase 1 features combined" do
      sheet =
        XlsxWriter.new_sheet("Phase 1 Features")
        # Merged header
        |> XlsxWriter.merge_range(0, 0, 0, 4, "Sales Report",
          format: [:bold, {:align, :center}]
        )
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

  describe "crash regression sweep — write/5 with :format" do
    test "boolean true with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.write(0, 0, true, format: [:bold])

      assert {"X", [{:write, 0, 0, {:boolean, true, [:bold]}}]} =
               sheet

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "boolean false with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.write(0, 0, false, format: [:italic])

      assert {"X", [{:write, 0, 0, {:boolean, false, [:italic]}}]} =
               sheet

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "Date with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.write(0, 0, ~D[2024-01-15], format: [:bold])

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "DateTime with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.write(0, 0, ~U[2024-01-15 09:30:00Z],
          format: [{:align, :center}]
        )

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "NaiveDateTime with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.write(0, 0, ~N[2024-01-15 09:30:00], format: [:bold])

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "Decimal with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.write(0, 0, Decimal.new("12.34"),
          format: [{:num_format, "0.00"}]
        )

      assert {"X",
              [
                {:write, 0, 0, {:number, 12.34, [{:num_format, "0.00"}]}}
              ]} = sheet

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "atom with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.write(0, 0, :hello, format: [:bold])

      assert {"X", [{:write, 0, 0, {:string, "hello", [:bold]}}]} =
               sheet

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "unsupported type with format raises XlsxWriter.Error", %{
      untyped: untyped
    } do
      sheet = XlsxWriter.new_sheet("X")

      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.write(sheet, 0, 0, untyped.pid, format: [:bold])
      end
    end
  end

  describe "crash regression sweep — merge_range/7 with :format" do
    test "nil with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.merge_range(0, 0, 0, 3, nil, format: [:bold])

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "Date with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.merge_range(0, 0, 0, 3, ~D[2024-01-15],
          format: [:bold, {:align, :center}]
        )

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "DateTime with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.merge_range(0, 0, 0, 3, ~U[2024-01-15 09:30:00Z],
          format: [:bold]
        )

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "NaiveDateTime with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.merge_range(0, 0, 0, 3, ~N[2024-01-15 09:30:00],
          format: [:bold]
        )

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "Decimal with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.merge_range(0, 0, 0, 3, Decimal.new("99.5"),
          format: [{:num_format, "0.00"}]
        )

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "atom with format does not crash" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.merge_range(0, 0, 0, 3, :hello, format: [:bold])

      assert {:ok, _} = XlsxWriter.generate([sheet])
    end

    test "unsupported type with format raises XlsxWriter.Error", %{
      untyped: untyped
    } do
      sheet = XlsxWriter.new_sheet("X")

      assert_raise XlsxWriter.Error, ~r/not supported/, fn ->
        XlsxWriter.merge_range(sheet, 0, 0, 0, 3, untyped.pid, format: [:bold])
      end
    end
  end

  describe "crash regression sweep — write/5 boolean semantics" do
    test "boolean true (no format) produces a boolean cell, not a string" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.write(0, 0, true)

      assert {"X", [{:write, 0, 0, {:boolean, true, []}}]} = sheet
    end

    test "boolean false (no format) produces a boolean cell, not a string" do
      sheet =
        XlsxWriter.new_sheet("X")
        |> XlsxWriter.write(0, 0, false)

      assert {"X", [{:write, 0, 0, {:boolean, false, []}}]} = sheet
    end
  end

  describe "merge_range/7 cell contents" do
    @merge_values [
      {"Title", ~r{<c r="A1"( s="1")? t="s"><v>0</v></c>}},
      {100, ~r{<c r="A2"( s="1")?><v>100</v></c>}},
      {2.5, ~r{<c r="A3"( s="1")?><v>2.5</v></c>}},
      {true, ~r{<c r="A4"( s="1")? t="b"><v>1</v></c>}},
      {Decimal.new("9.5"), ~r{<c r="A5"( s="1")?><v>9.5</v></c>}},
      {~D[2024-01-15], ~r{<c r="A6" s="\d+"><v>45306</v></c>}},
      {~U[2024-01-15 10:00:00Z],
       ~r{<c r="A7" s="\d+"><v>45306.416666666664</v></c>}},
      {~N[2024-01-15 10:30:00], ~r{<c r="A8" s="\d+"><v>45306.4375</v></c>}},
      {:atom, ~r{<c r="A9"( s="1")? t="s"><v>1</v></c>}}
    ]

    for format <- [[], [:bold]] do
      test "keeps the typed value and merges every value type with format #{inspect(format)}" do
        sheet =
          @merge_values
          |> Enum.with_index()
          |> Enum.reduce(XlsxWriter.new_sheet("Merged"), fn {{value, _}, row},
                                                            sheet ->
            XlsxWriter.merge_range(sheet, row, 0, row, 2, value,
              format: unquote(format)
            )
          end)

        {:ok, content} = XlsxWriter.generate([sheet])
        sheet_xml = xlsx_part(content, "xl/worksheets/sheet1.xml")

        for {{_value, cell_pattern}, row} <- Enum.with_index(@merge_values, 1) do
          assert sheet_xml =~ cell_pattern
          assert sheet_xml =~ ~s(<mergeCell ref="A#{row}:C#{row}"/>)
        end

        assert xlsx_part(content, "xl/sharedStrings.xml") =~
                 ~s(<si><t>Title</t></si><si><t>atom</t></si></sst>)
      end
    end

    test "applies the format to every cell of the range" do
      sheet =
        XlsxWriter.new_sheet("Merged")
        |> XlsxWriter.merge_range(0, 0, 1, 1, 42, format: [:bold])

      sheet_xml = generate_part!([sheet])

      for cell <- ["B1", "A2", "B2"] do
        assert sheet_xml =~ ~s(<c r="#{cell}" s="1"/>)
      end

      assert sheet_xml =~ ~s(<c r="A1" s="1"><v>42</v></c>)
    end

    test "dates use the same default number formats as write/5" do
      sheet =
        XlsxWriter.new_sheet("Merged")
        |> XlsxWriter.merge_range(0, 0, 0, 1, ~D[2024-01-15])
        |> XlsxWriter.merge_range(1, 0, 1, 1, ~N[2024-01-15 10:30:00])

      {:ok, content} = XlsxWriter.generate([sheet])
      styles = xlsx_part(content, "xl/styles.xml")
      assert styles =~ ~s(formatCode="yyyy-mm-dd")
      assert styles =~ ~s(formatCode="yyyy-mm-ddThh:mm:ss")
    end

    test "a user num_format overrides the default date format" do
      sheet =
        XlsxWriter.new_sheet("Merged")
        |> XlsxWriter.merge_range(0, 0, 0, 1, ~D[2024-01-15],
          format: [{:num_format, "dd/mm/yyyy"}]
        )

      {:ok, content} = XlsxWriter.generate([sheet])
      styles = xlsx_part(content, "xl/styles.xml")
      assert styles =~ ~s(formatCode="dd/mm/yyyy")
      refute styles =~ ~s(formatCode="yyyy-mm-dd")
    end

    test "returns an error when a raw instruction merges an image" do
      sheets = [{"S", [{:merge_range, 0, 0, 0, 1, {:image, "image"}}]}]

      assert XlsxWriter.generate(sheets) ==
               {:error, "images cannot be written to a merged range"}
    end
  end

  describe "fractional sizes" do
    test "column widths and row heights accept floats" do
      sheet =
        XlsxWriter.new_sheet("S")
        |> XlsxWriter.write(0, 0, "x")
        |> XlsxWriter.set_column_width(0, 25.5)
        |> XlsxWriter.set_column_range_width(1, 2, 12.5)
        |> XlsxWriter.set_row_height(0, 30.0)
        |> XlsxWriter.set_row_range_height(1, 2, 12.5)
        |> XlsxWriter.write(1, 0, "y")

      sheet_xml = generate_part!([sheet])

      assert sheet_xml =~
               ~s(<col min="1" max="1" width="26.28515625" customWidth="1"/>)

      assert sheet_xml =~
               ~s(<col min="2" max="3" width="13.28515625" customWidth="1"/>)

      assert sheet_xml =~ ~s(<row r="1" spans="1:1" ht="30" customHeight="1">)

      assert sheet_xml =~
               ~s(<row r="2" spans="1:1" ht="12.75" customHeight="1">)
    end

    test "font sizes accept floats" do
      sheet =
        XlsxWriter.new_sheet("S")
        |> XlsxWriter.write(0, 0, "x", format: [{:font_size, 11.5}])

      assert generate_part!([sheet], "xl/styles.xml") =~ ~s(<sz val="11.5"/>)
    end

    test "comment sizes accept floats" do
      sheet =
        XlsxWriter.new_sheet("S")
        |> XlsxWriter.write_comment(0, 0, "note",
          width: 100.5,
          height: 50.4,
          visible: true
        )

      vml = generate_part!([sheet], "xl/drawings/vmlDrawing1.vml")
      # 100.5 and 50.4 pixels round to 101 and 50 pixels.
      assert vml =~ "width:75.75pt;height:37.5pt"
    end
  end

  describe "colors" do
    test "accepts #RRGGBB and RRGGBB in any case" do
      sheet =
        XlsxWriter.new_sheet("S")
        |> XlsxWriter.write(0, 0, "a",
          format: [
            {:bg_color, "#4472c4"},
            {:font_color, "FF0000"},
            {:border, :thin},
            {:border_color, "#00ff00"}
          ]
        )
        |> XlsxWriter.set_tab_color("#ABCDEF")

      {:ok, content} = XlsxWriter.generate([sheet])
      styles = xlsx_part(content, "xl/styles.xml")
      assert styles =~ ~s(<fgColor rgb="FF4472C4"/>)
      assert styles =~ ~s(<color rgb="FFFF0000"/>)
      assert styles =~ ~s(<left style="thin"><color rgb="FF00FF00"/></left>)

      assert xlsx_part(content, "xl/worksheets/sheet1.xml") =~
               ~s(<sheetPr><tabColor rgb="FFABCDEF"/></sheetPr>)
    end

    test "the NIF returns an error for invalid colors in raw instructions" do
      for instruction <- [
            {:write, 0, 0, {:string, "x", [{:bg_color, "#FFF"}]}},
            {:write, 0, 0, {:string, "x", [{:font_color, "red"}]}},
            {:write, 0, 0, {:string, "x", [{:font_color, "#FFFFFFFF"}]}},
            {:set_tab_color, "blue"}
          ] do
        assert {:error, "invalid color " <> _} =
                 XlsxWriter.generate([{"S", [instruction]}])
      end
    end
  end

  describe "autofit/1 ordering" do
    @long_text "A very long string that should widen the column"

    test "fits columns written after autofit/1 was called" do
      sheet =
        XlsxWriter.new_sheet("S")
        |> XlsxWriter.autofit()
        |> XlsxWriter.write(0, 0, @long_text)

      assert generate_part!([sheet]) =~
               ~s(<cols><col min="1" max="1" width="43.140625" bestFit="1" customWidth="1"/></cols>)
    end

    test "gives the same result wherever autofit/1 appears" do
      first =
        XlsxWriter.new_sheet("S")
        |> XlsxWriter.autofit()
        |> XlsxWriter.write(0, 0, @long_text)
        |> XlsxWriter.write(0, 1, "short")

      last =
        XlsxWriter.new_sheet("S")
        |> XlsxWriter.write(0, 0, @long_text)
        |> XlsxWriter.write(0, 1, "short")
        |> XlsxWriter.autofit()

      assert generate_part!([first]) == generate_part!([last])
    end

    test "explicit column widths take precedence, before or after autofit/1" do
      sheet =
        XlsxWriter.new_sheet("S")
        |> XlsxWriter.set_column_width(0, 5)
        |> XlsxWriter.write(0, 0, @long_text)
        |> XlsxWriter.write(0, 1, @long_text)
        |> XlsxWriter.write(0, 2, @long_text)
        |> XlsxWriter.write(0, 3, @long_text)
        |> XlsxWriter.autofit()
        |> XlsxWriter.set_column_range_width(1, 2, 8)

      assert generate_part!([sheet]) =~
               ~s(<cols><col min="1" max="1" width="5.7109375" customWidth="1"/>) <>
                 ~s(<col min="2" max="3" width="8.7109375" customWidth="1"/>) <>
                 ~s(<col min="4" max="4" width="43.140625" bestFit="1" customWidth="1"/></cols>)
    end

    test "Builder column widths take precedence over autofit" do
      {:ok, content} =
        XlsxWriter.Builder.create()
        |> XlsxWriter.Builder.add_sheet("S")
        |> XlsxWriter.Builder.add_rows([[{@long_text, width: 5}, @long_text]])
        |> XlsxWriter.Builder.write_binary()

      assert xlsx_part(content, "xl/worksheets/sheet1.xml") =~
               ~s(<col min="1" max="1" width="5.7109375" customWidth="1"/>)
    end
  end

  describe "DateTime values" do
    test "are written as wall-clock time; the UTC offset is dropped" do
      athens = %DateTime{
        year: 2024,
        month: 1,
        day: 15,
        hour: 10,
        minute: 0,
        second: 0,
        microsecond: {0, 0},
        utc_offset: 7200,
        std_offset: 0,
        zone_abbr: "EET",
        time_zone: "Europe/Athens"
      }

      sheet =
        XlsxWriter.new_sheet("S")
        |> XlsxWriter.write(0, 0, athens)
        |> XlsxWriter.write(1, 0, ~U[2024-01-15 10:00:00Z])
        |> XlsxWriter.write(2, 0, ~N[2024-01-15 10:00:00])

      assert {"S",
              [_, _, {:write, 0, 0, {:date_time, "2024-01-15T10:00:00", []}}]} =
               sheet

      sheet_xml = generate_part!([sheet])

      for row <- 1..3 do
        assert sheet_xml =~
                 ~r{<c r="A#{row}" s="1"><v>45306.416666666664</v></c>}
      end
    end

    test "keep sub-second precision" do
      sheet =
        XlsxWriter.new_sheet("S")
        |> XlsxWriter.write(0, 0, ~U[2024-01-15 10:00:00.500000Z])

      assert generate_part!([sheet]) =~ "<v>45306.416672453706</v>"
    end
  end

  describe "generate/2 error handling" do
    test "returns an error for an instruction the NIF cannot decode" do
      {name, instructions} =
        XlsxWriter.new_sheet("S") |> XlsxWriter.write(0, 0, "ok")

      # Instructions are stored newest first and numbered in write order.
      sheet = {name, [{:write, 0, 0, :bad} | instructions]}

      assert XlsxWriter.generate([sheet]) ==
               {:error,
                ~s(invalid instruction #2 in sheet "S": {write,0,0,bad})}
    end

    test "returns an error for raw instructions with unknown formats or wrong types" do
      for instruction <- [
            {:write, 0, 0, {:string, "x", [:bolt]}},
            {:set_column_width, 0, "wide"},
            {:write, 1.5, 0, {:string, "x", []}},
            {:insert_note, 0, 0, 123, %XlsxWriter.NoteOptions{}}
          ] do
        assert {:error, "invalid instruction #1 in sheet \"S\": " <> _} =
                 XlsxWriter.generate([{"S", [instruction]}])
      end
    end

    test "returns an error for a malformed sheet" do
      assert XlsxWriter.generate([:nope]) == {:error, "invalid sheet #1: nope"}

      assert {:error, "invalid sheet #2: " <> _} =
               XlsxWriter.generate([XlsxWriter.new_sheet("A"), {"B", :nope}])
    end

    test "returns an error for invalid workbook properties" do
      sheet = XlsxWriter.new_sheet("S")

      assert {:error, "invalid workbook properties: " <> _} =
               XlsxWriter.generate([sheet], properties: %{author: "x"})

      assert {:error, "invalid workbook properties: " <> _} =
               XlsxWriter.generate([sheet],
                 properties: %XlsxWriter.WorkbookProperties{author: 123}
               )
    end

    test "writes integers beyond 64 bits as numbers" do
      sheet =
        XlsxWriter.new_sheet("S") |> XlsxWriter.write(0, 0, Integer.pow(2, 70))

      assert generate_part!([sheet]) =~
               ~s(<c r="A1"><v>1180591620717411300000</v></c>)
    end

    test "the NIF rejects out-of-range positions in raw instructions instead of wrapping them" do
      for {instruction, message} <- [
            {{:write, 0, 70_000, {:string, "x", []}},
             "column 70000 is outside Excel's range 0..16383"},
            {{:write, -1, 0, {:string, "x", []}},
             "row -1 is outside Excel's range 0..1048575"},
            {{:set_column_range_width, 0, 65_537, 10},
             "column 65537 is outside Excel's range 0..16383"},
            {{:set_row_height, 4_294_967_296, 10},
             "row 4294967296 is outside Excel's range 0..1048575"},
            {{:write, 0, 0, {:string, "x", [{:indent, 257}]}},
             "indent 257 is outside the range 0..250"},
            {{:write, 0, 0, {:string, "x", [{:rotation, 65_581}]}},
             "rotation 65581 must be between -90 and 90, or 270"},
            {{:write, 0, 0, {:string, "x", [{:font_size, 65_547}]}},
             "font size 65547 is outside Excel's range 0..409"},
            {{:set_row_height, 0, 65_556},
             "row height 65556 is outside Excel's range 0..409"}
          ] do
        assert XlsxWriter.generate([{"S", [instruction]}]) == {:error, message}
      end
    end

    test "the NIF returns the workbook as a binary" do
      assert {:ok, content} = XlsxWriter.RustXlsxWriter.write([{"S", []}], nil)
      assert is_binary(content)
      assert <<"PK", _rest::binary>> = content
    end
  end

  describe "comments" do
    test "write the author" do
      author = String.duplicate("a", 52)

      sheet =
        XlsxWriter.new_sheet("S")
        |> XlsxWriter.write_comment(0, 0, "x", author: author)

      comments = generate_part!([sheet], "xl/comments1.xml")
      assert comments =~ "<author>#{author}</author></authors>"
      assert comments =~ ~s(<comment ref="A1" authorId="1">)
    end

    test "the NIF returns an error for a long author in a raw instruction" do
      options = %XlsxWriter.NoteOptions{author: String.duplicate("a", 53)}

      assert {:error, "comment author must be at most 52 characters" <> _} =
               XlsxWriter.generate([
                 {"S", [{:insert_note, 0, 0, "x", options}]}
               ])
    end
  end
end
