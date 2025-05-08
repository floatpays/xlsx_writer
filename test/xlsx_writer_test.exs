defmodule XlsxWriterTest do
  use ExUnit.Case

  describe "write/1" do
    test "write xlsx file" do
      assert XlsxWriter.RustXlsxWriter.test_new_workbook() == {:ok, {}}
    end
  end
end
