defmodule XlsxWriterTest do
  use ExUnit.Case

  alias XlsxWriter.Workbook

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

      assert {:ok, content} = XlsxWriter.Workbook.generate(sheets)

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

      assert {:ok, content} = XlsxWriter.Workbook.generate(sheets)

      File.write!("test2.xlsx", content)
    end

    test "write xlsx file with porcelain" do
      filename = "test2.xlsx"

      sheet1 =
        Workbook.new_sheet("sheet number one")
        |> Workbook.write(0, 0, "col1", format: [:bold])
        |> Workbook.write(0, 1, "col2", format: [:bold, {:align, :center}])
        |> Workbook.write(0, 2, "col3", format: [:bold, {:align, :right}])
        |> Workbook.write(0, 3, nil)
        |> Workbook.set_column_width(0, 40)
        |> Workbook.set_column_width(3, 60)
        |> Workbook.write(1, 0, "row 2 col 1")
        |> Workbook.write(1, 1, 1.0)
        |> Workbook.write_formula(1, 2, "=B2 + 2")
        |> Workbook.write_formula(2, 1, "=PI()")
        |> Workbook.write_image(3, 0, File.read!("bird.jpeg"))
        |> Workbook.write(4, 3, 1)
        |> Workbook.write(5, 3, DateTime.utc_now())
        |> Workbook.write(6, 3, NaiveDateTime.utc_now())
        |> Workbook.write(7, 3, Date.utc_today())
        |> Workbook.write(8, 3, Decimal.new("20.12"))

      sheet2 =
        Workbook.new_sheet("sheet number two")
        |> Workbook.write(0, 0, "col1")

      {:ok, content} = Workbook.generate([sheet1, sheet2])

      File.write!(filename, content)
    end

    test "write xlsx file with numeric format" do
      filename = "test2.xlsx"

      sheet1 =
        Workbook.new_sheet("sheet number one")
        |> Workbook.write(0, 0, 999.99, format: [
          {:num_format, "[$R] #,##0.00"}
        ])
        |> Workbook.write(1, 0, 888, format: [{:num_format, "0,000.00"}])

      {:ok, content} = Workbook.generate([sheet1])

      File.write!(filename, content)
    end

    test "write xlsx file with unsupported format" do
      assert_raise XlsxWriter.Error, fn ->
        Workbook.new_sheet("sheet number one")
        |> Workbook.write(0, 0, self())
      end
    end
  end
end
