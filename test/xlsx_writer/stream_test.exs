defmodule XlsxWriter.StreamTest do
  use ExUnit.Case, async: true

  alias XlsxWriter.Stream, as: XlsxStream

  @tag :tmp_dir
  test "writes enumerable rows across multiple sheets without collecting a workbook",
       %{tmp_dir: dir} do
    path = Path.join(dir, "large.xlsx")
    rows = Elixir.Stream.map(1..1_200, fn n -> [n, "Row #{n}"] end)

    assert :ok =
             XlsxStream.write_file(
               path,
               [
                 {"Data", rows},
                 {"Summary",
                  [[{"Total", format: [:bold]}, {:formula, "=1000+200"}]]}
               ],
               batch_size: 37,
               tempdir: dir,
               properties: %XlsxWriter.WorkbookProperties{author: "Stream test"}
             )

    assert {:ok, entries} = :zip.extract(String.to_charlist(path), [:memory])

    sheet1 =
      entries
      |> Enum.find(fn {name, _} -> name == ~c"xl/worksheets/sheet1.xml" end)
      |> elem(1)

    sheet2 =
      entries
      |> Enum.find(fn {name, _} -> name == ~c"xl/worksheets/sheet2.xml" end)
      |> elem(1)

    styles =
      entries
      |> Enum.find(fn {name, _} -> name == ~c"xl/styles.xml" end)
      |> elem(1)

    assert sheet1 =~ "<row r=\"1200\""
    assert sheet1 =~ "Row 1200"
    assert sheet2 =~ "Total"
    assert sheet2 =~ "<f>1000+200</f>"
    assert styles =~ "<b/>"
    assert Enum.sort(File.ls!(dir)) == ["large.xlsx"]
  end

  @tag :tmp_dir
  test "supports bounded calls and forward row skips", %{tmp_dir: dir} do
    path = Path.join(dir, "skipped.xlsx")
    assert {:ok, writer} = XlsxStream.open(path, tempdir: dir)
    assert {:ok, writer} = XlsxStream.add_sheet(writer, "Data")

    setup =
      XlsxWriter.new_sheet("Data")
      |> XlsxWriter.set_column_width(2, 25)
      |> XlsxWriter.freeze_panes(1, 0)

    assert {:ok, writer} = XlsxStream.configure_sheet(writer, setup)
    assert {:ok, writer} = XlsxStream.write_rows(writer, [["First"]])

    assert {:ok, writer} =
             XlsxStream.write_rows(writer, [["Later"]],
               start_row: 4,
               start_col: 2
             )

    assert :ok = XlsxStream.finish(writer)
    assert {:error, "Workbook is closed"} = XlsxStream.finish(writer)

    assert {:ok, entries} = :zip.extract(String.to_charlist(path), [:memory])

    sheet =
      entries
      |> Enum.find(fn {name, _} -> name == ~c"xl/worksheets/sheet1.xml" end)
      |> elem(1)

    assert sheet =~ "<c r=\"C5\""
    assert sheet =~ ~s(<col min="3" max="3" width=")
    assert sheet =~ "state=\"frozen\""
  end

  @tag :tmp_dir
  test "rejects backwards rows and abort removes temporary files", %{
    tmp_dir: dir
  } do
    path = Path.join(dir, "aborted.xlsx")
    assert {:ok, writer} = XlsxStream.open(path, tempdir: dir)
    assert {:ok, writer} = XlsxStream.add_sheet(writer, "Data")
    assert {:ok, writer} = XlsxStream.write_rows(writer, [[1]], start_row: 5)
    assert {:error, _} = XlsxStream.write_rows(writer, [[2]], start_row: 4)
    assert :ok = XlsxStream.abort(writer)
    refute File.exists?(path)
    assert File.ls!(dir) == []
  end

  @tag :tmp_dir
  test "aborts when an input stream raises", %{tmp_dir: dir} do
    path = Path.join(dir, "failed.xlsx")

    rows =
      Elixir.Stream.map(1..3, fn
        3 -> raise "source failed"
        n -> [n]
      end)

    assert_raise RuntimeError, "source failed", fn ->
      XlsxStream.write_file(path, [{"Data", rows}], batch_size: 1, tempdir: dir)
    end

    assert File.ls!(dir) == []
  end
end
