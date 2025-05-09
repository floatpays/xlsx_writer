defmodule XlsxWriterTest do
  use ExUnit.Case

  describe "write/1" do
    test "write xlsx file" do
      bird_content = File.read!("bird.jpeg")

      data = [
        {0, 0, {:string, "foo"}},
        {0, 1, {:string, "bar"}},
        {1, 1, {:image_path, "bird.jpeg"}},
        {1, 2, {:image, bird_content |> :binary.bin_to_list()}}
      ]

      assert {:ok, _} = XlsxWriter.write(data, "foo.xlsx")
    end
  end
end
