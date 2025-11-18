defmodule XlsxWriter.CommentTest do
  use ExUnit.Case, async: true

  alias XlsxWriter

  describe "write_comment/5" do
    test "adds simple comment to cell" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_comment(0, 0, "This is a comment")

      assert {"Test", [{:insert_note, 0, 0, "This is a comment", options}]} = sheet
      assert %XlsxWriter.NoteOptions{} = options
      assert options.author == nil
      assert options.visible == nil
      assert options.width == nil
      assert options.height == nil
    end

    test "adds comment with author" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_comment(0, 0, "Review this value", author: "John Doe")

      assert {"Test", [{:insert_note, 0, 0, "Review this value", options}]} = sheet
      assert options.author == "John Doe"
    end

    test "adds visible comment" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_comment(0, 0, "Important note", visible: true)

      assert {"Test", [{:insert_note, 0, 0, "Important note", options}]} = sheet
      assert options.visible == true
    end

    test "adds comment with custom dimensions" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_comment(0, 0, "Large comment", width: 300, height: 200)

      assert {"Test", [{:insert_note, 0, 0, "Large comment", options}]} = sheet
      assert options.width == 300
      assert options.height == 200
    end

    test "adds comment with all options" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_comment(0, 0, "Full featured note",
          author: "Jane Smith",
          visible: true,
          width: 250,
          height: 150
        )

      assert {"Test", [{:insert_note, 0, 0, "Full featured note", options}]} = sheet
      assert options.author == "Jane Smith"
      assert options.visible == true
      assert options.width == 250
      assert options.height == 150
    end

    test "validates cell position" do
      sheet = XlsxWriter.new_sheet("Test")

      assert_raise ArgumentError, ~r/Row index must be/, fn ->
        XlsxWriter.write_comment(sheet, -1, 0, "Comment")
      end

      assert_raise ArgumentError, ~r/Column index must be/, fn ->
        XlsxWriter.write_comment(sheet, 0, -1, "Comment")
      end
    end

    test "validates text is a string" do
      sheet = XlsxWriter.new_sheet("Test")

      assert_raise ArgumentError, ~r/Comment text must be a string/, fn ->
        XlsxWriter.write_comment(sheet, 0, 0, 123)
      end

      assert_raise ArgumentError, ~r/Comment text must be a string/, fn ->
        XlsxWriter.write_comment(sheet, 0, 0, nil)
      end
    end

    test "comments can be added to cells with data" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Cell value")
        |> XlsxWriter.write_comment(0, 0, "Explanation of value")

      {"Test", instructions} = sheet
      assert length(instructions) == 2
      assert {:insert_note, 0, 0, "Explanation of value", _} = Enum.at(instructions, 0)
      assert {:write, 0, 0, {:string, "Cell value"}} = Enum.at(instructions, 1)
    end

    test "multiple comments can be added to different cells" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write_comment(0, 0, "Comment A")
        |> XlsxWriter.write_comment(1, 0, "Comment B")
        |> XlsxWriter.write_comment(0, 1, "Comment C")

      {"Test", instructions} = sheet
      assert length(instructions) == 3
      assert {:insert_note, 0, 1, "Comment C", _} = Enum.at(instructions, 0)
      assert {:insert_note, 1, 0, "Comment B", _} = Enum.at(instructions, 1)
      assert {:insert_note, 0, 0, "Comment A", _} = Enum.at(instructions, 2)
    end
  end

  describe "integration tests" do
    test "generates valid xlsx with comments" do
      sheet =
        XlsxWriter.new_sheet("Comments")
        |> XlsxWriter.write(0, 0, "Value", format: [:bold])
        |> XlsxWriter.write_comment(0, 0, "This is a test comment")
        |> XlsxWriter.write(1, 0, 100)
        |> XlsxWriter.write_comment(1, 0, "Important value", author: "Reviewer")

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert is_binary(content)
      # XLSX files start with PK (ZIP magic bytes)
      assert binary_part(content, 0, 2) == "PK"
    end

    test "generates xlsx with visible comment" do
      sheet =
        XlsxWriter.new_sheet("Test")
        |> XlsxWriter.write(0, 0, "Note this")
        |> XlsxWriter.write_comment(0, 0, "Always visible note",
          visible: true,
          width: 300,
          height: 200
        )

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert is_binary(content)
      assert binary_part(content, 0, 2) == "PK"
    end

    test "generates xlsx with comments and complex formatting" do
      sheet =
        XlsxWriter.new_sheet("Dashboard")
        |> XlsxWriter.write(0, 0, "Metric", format: [:bold, {:bg_color, "#4472C4"}])
        |> XlsxWriter.write_comment(0, 0, "Key performance indicator")
        |> XlsxWriter.write(0, 1, "Value", format: [:bold, {:bg_color, "#4472C4"}])
        |> XlsxWriter.write(1, 0, "Sales")
        |> XlsxWriter.write(1, 1, 1500, format: [{:num_format, "$#,##0.00"}])
        |> XlsxWriter.write_comment(1, 1, "Q1 target: $2000", author: "Manager", visible: true)

      assert {:ok, content} = XlsxWriter.generate([sheet])
      assert is_binary(content)
    end
  end
end
