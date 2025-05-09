defmodule XlsxWriterTest do
  use ExUnit.Case

  describe "write/1" do
    test "write xlsx file" do
      bird_content = File.read!("bird.jpeg")

      data = [
        {:insert, 0, 0, {:string, "this is new"}},
        {:insert, 0, 1, {:float, 12.12}},
        {:insert, 0, 3, {:image_path, "bird.jpeg"}},
        {:insert, 1, 2, {:image, bird_content}},
        {:set_column_width, 0, 30},
        {:set_row_height, 0, 30}
      ]

      assert {:ok, _} = XlsxWriter.write(data, "foo.xlsx")
    end
  end
end
