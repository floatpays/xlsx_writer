defmodule XlsxWriter.ValidationTest do
  use ExUnit.Case, async: true

  alias XlsxWriter.Validation

  doctest XlsxWriter.Validation

  describe "validate_cell_position!/2" do
    test "accepts the corners of Excel's grid" do
      assert Validation.validate_cell_position!(0, 0) == :ok
      assert Validation.validate_cell_position!(1_048_575, 16_383) == :ok
    end

    test "rejects non-integer and out-of-range indices" do
      for {row, col, message} <- [
            {1_048_576, 0,
             "Row index must be at most 1048575 (Excel's last row), got: 1048576"},
            {0, 16_384,
             "Column index must be at most 16383 (Excel's last column), got: 16384"},
            {-1, 0, "Row index must be non-negative, got: -1"},
            {0, -1, "Column index must be non-negative, got: -1"},
            {"a", 0, ~s(Row index must be an integer, got: "a")},
            {0, nil, "Column index must be an integer, got: nil"}
          ] do
        assert_raise ArgumentError, message, fn ->
          Validation.validate_cell_position!(row, col)
        end
      end
    end
  end

  describe "validate_color!/1" do
    test "accepts six hex digits with or without #, in any case" do
      for color <- ["#FF0000", "#ff00aa", "4472C4", "abcdef"] do
        assert Validation.validate_color!(color) == :ok
      end
    end

    test "rejects anything else" do
      for color <- ["#FFF", "red", "#FFFFFFFF", "", "#GGGGGG", " #FF0000", nil] do
        assert_raise ArgumentError, ~r/Expected a string hex color/, fn ->
          Validation.validate_color!(color)
        end
      end
    end
  end

  describe "validate_formats!/1" do
    test "accepts every documented option" do
      formats = [
        :bold,
        :italic,
        :strikethrough,
        :superscript,
        :subscript,
        :text_wrap,
        :shrink,
        {:rotation, -90},
        {:rotation, 270},
        {:indent, 0},
        {:align, :left},
        {:valign, :distributed},
        {:num_format, "0.00"},
        {:bg_color, "#FFFF00"},
        {:font_color, "#FF0000"},
        {:font_size, 10.5},
        {:font_name, "Arial"},
        {:underline, :double_accounting},
        {:pattern, :gray0625},
        {:border, :slant_dash_dot},
        {:border_top, :thin},
        {:border_bottom, :thin},
        {:border_left, :thin},
        {:border_right, :thin},
        {:border_color, "#000000"},
        {:border_top_color, "#000000"},
        {:border_bottom_color, "#000000"},
        {:border_left_color, "#000000"},
        {:border_right_color, "#000000"}
      ]

      assert Validation.validate_formats!(formats) == :ok
    end
  end

  describe "validate_comment_options!/1" do
    test "rejects invalid options" do
      for {options, message} <- [
            {[author: String.duplicate("a", 53)], ~r/at most 52 characters/},
            {[author: :me], ~r/Comment author must be a string/},
            {[visible: "yes"], ~r/Comment :visible must be a boolean/},
            {[width: 0.4], ~r/Comment :width must be a number of pixels/}
          ] do
        assert_raise ArgumentError, message, fn ->
          Validation.validate_comment_options!(options)
        end
      end
    end

    test "counts the author length in code points, like rust_xlsxwriter" do
      # Precomposed "é" is one code point: 52 of them fit.
      assert Validation.validate_comment_options!(
               author: String.duplicate("é", 52)
             ) ==
               :ok

      # "e" plus a combining accent is one grapheme but two code points, so
      # 30 of them are 60 code points: rust_xlsxwriter would drop the author.
      combining_author = String.duplicate("é", 30)
      assert String.length(combining_author) == 30

      assert_raise ArgumentError, ~r/at most 52 characters/, fn ->
        Validation.validate_comment_options!(author: combining_author)
      end
    end

    test "rounds comment sizes before checking them" do
      assert Validation.validate_comment_options!(width: 0.6, height: 1) == :ok

      assert_raise ArgumentError, ~r/Comment :height/, fn ->
        Validation.validate_comment_options!(height: 0.4)
      end
    end
  end
end
