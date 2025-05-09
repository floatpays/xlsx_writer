defmodule XlsxWriterTest do
  use ExUnit.Case

  describe "write/1" do
    test "write xlsx file" do
      bird_content = File.read!("bird.jpeg")

      data = [
        {:write, 0, 0, {:string, "this is new"}},
        # {:write_with_format, 0, 0, {:string, "this is new"}, [:bold]},
        {:write, 0, 1, {:float, 12.12}},
        {:write, 0, 3, {:image_path, "bird.jpeg"}},
        {:write, 1, 2, {:image, bird_content}},
        {:write, 2, 0, {:date, 2024, 10, 10}},
        {:set_column_width, 0, 30},
        {:set_row_height, 0, 30}
      ]

      assert {:ok, _} = XlsxWriter.write(data, "foo.xlsx")
    end
  end
end
