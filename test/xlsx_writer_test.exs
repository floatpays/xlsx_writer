defmodule XlsxWriterTest do
  use ExUnit.Case

  describe "write/1" do
    test "write xlsx file" do
      data = [
        {0, 0, {:string, "foo"}},
        {0, 1, {:string, "bar"}},
        {1, 1, {:image, "bird.jpeg"}}
      ]

      assert {:ok, _} = XlsxWriter.write(data, "foo.xlsx")
    end
  end
end
