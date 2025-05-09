defmodule XlsxWriterTest do
  use ExUnit.Case

  describe "write/1" do
    test "write xlsx file" do
      bird_content = File.read!("bird.jpeg")

      data = [
        {:write, 9, 0,
         {:string_with_format, "this is new", [{:align, :right}]}},
        {:write, 0, 0,
         {:string_with_format, "this is new", [:bold, {:align, :center}]}},
        {:write, 0, 1, {:float, 12.12}},
        {:write, 0, 3, {:image_path, "bird.jpeg"}},
        {:write, 1, 2, {:image, bird_content}},
        {:write, 2, 0, {:date, 2024, 10, 10}},
        {:set_column_width, 0, 30},
        {:set_row_height, 0, 30}
      ]

      assert {:ok, _} = XlsxWriter.write(data, "foo.xlsx")
    end

    # test "write xlsx file" do
    #   Workbook.new()
    #   |> Workbook.write(9, 0, "foo")
    #   |> Workbook.set_colum_width(9, 30)
    #   |> Workbook.write!()
    # end
  end
end
