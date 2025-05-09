defmodule XlsxWriterTest do
  use ExUnit.Case

  describe "write/1" do
    test "write xlsx file" do
      bird_content = File.read!("bird.jpeg")

      data = [
        {0, 0, {:string, "foo"}},
        {10, 0, {:float, 10.12}},
        {0, 1, {:string, "bar"}},
        {1, 1, {:image_path, "bird.jpeg"}},
        {1, 2, {:image, bird_content}},
        # TOOD: we only need the column here, not the row
        # how to do this?
        {0, 0, {:column_width, 30}},
        {0, 0, {:row_height, 30}}
      ]

      assert {:ok, _} = XlsxWriter.write(data, "foo.xlsx")
    end
  end
end
