defmodule XlsxWriterTest do
  use ExUnit.Case

  describe "write/1" do
    test "write xlsx file" do
      data = [
        {0, 0, "zero zero"},
        {0, 1, "zero one"}
      ]

      assert {:ok, _} = XlsxWriter.write(data, "foo.xlsx")
    end
  end
end
