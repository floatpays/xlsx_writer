defmodule XlsxWriter.Validation do
  @moduledoc """
  Validation functions for XlsxWriter inputs.

  `XlsxWriter` and `XlsxWriter.Builder` call these when an instruction is
  added, so that mistakes fail early, with a clear message, instead of
  surfacing later from `XlsxWriter.generate/2` or being silently dropped.

  Every validation failure raises `ArgumentError`, except for the deprecated
  `validate_supported_type!/1`, which raises `XlsxWriter.Error` as before.
  """

  @max_row 1_048_575
  @max_column 16_383
  # The range of rust_xlsxwriter's indent (a u8); larger values used to wrap.
  @max_indent 255
  @max_comment_author_length 52
  # Comment sizes are whole pixels stored as u32 by rust_xlsxwriter.
  @max_comment_size 4_294_967_295

  @flag_formats [
    :bold,
    :italic,
    :strikethrough,
    :superscript,
    :subscript,
    :text_wrap,
    :shrink
  ]

  @color_formats [
    :bg_color,
    :font_color,
    :border_color,
    :border_top_color,
    :border_bottom_color,
    :border_left_color,
    :border_right_color
  ]

  @border_formats [
    :border,
    :border_top,
    :border_bottom,
    :border_left,
    :border_right
  ]

  @border_styles [
    :thin,
    :medium,
    :thick,
    :dashed,
    :dotted,
    :double,
    :hair,
    :medium_dashed,
    :dash_dot,
    :medium_dash_dot,
    :dash_dot_dot,
    :medium_dash_dot_dot,
    :slant_dash_dot
  ]

  @alignments [:left, :center, :right]
  @vertical_alignments [:top, :center, :bottom, :justify, :distributed]
  @patterns [:solid, :none, :gray125, :gray0625]
  @underline_styles [:single, :double, :single_accounting, :double_accounting]

  @doc """
  Validates a cell position: both indices must be integers within Excel's
  limits (rows `0..#{@max_row}`, columns `0..#{@max_column}`).

  ## Raises

  - `ArgumentError` if the row or column is not an integer or is out of range

  ## Examples

      iex> XlsxWriter.Validation.validate_cell_position!(0, 0)
      :ok

      iex> XlsxWriter.Validation.validate_cell_position!(-1, 0)
      ** (ArgumentError) Row index must be non-negative, got: -1

      iex> XlsxWriter.Validation.validate_cell_position!(0, 16_384)
      ** (ArgumentError) Column index must be at most 16383 (Excel's last column), got: 16384

  """
  def validate_cell_position!(row, col) do
    validate_row!(row)
    validate_column!(col)
  end

  # The validators below are internal: XlsxWriter and XlsxWriter.Builder call
  # them when an instruction is added.

  @doc false
  def validate_row!(row), do: validate_index!(row, "Row", @max_row, "row")

  @doc false
  def validate_column!(col),
    do: validate_index!(col, "Column", @max_column, "column")

  @doc false
  def validate_range!(first_row, first_col, last_row, last_col) do
    validate_row_range!(first_row, last_row)
    validate_column_range!(first_col, last_col)
  end

  @doc false
  def validate_row_range!(first_row, last_row) do
    validate_row!(first_row)
    validate_row!(last_row)
    validate_order!(first_row, last_row, "row")
  end

  @doc false
  def validate_column_range!(first_col, last_col) do
    validate_column!(first_col)
    validate_column!(last_col)
    validate_order!(first_col, last_col, "column")
  end

  @doc """
  Validates that image data is a non-empty binary.

  ## Raises

  - `ArgumentError` if the data is not a binary or is empty

  ## Examples

      iex> XlsxWriter.Validation.validate_image_binary!(<<1, 2, 3>>)
      :ok

      iex> XlsxWriter.Validation.validate_image_binary!(<<>>)
      ** (ArgumentError) Image binary cannot be empty

  """
  def validate_image_binary!(<<>>),
    do: raise(ArgumentError, "Image binary cannot be empty")

  def validate_image_binary!(image_binary) when is_binary(image_binary), do: :ok

  def validate_image_binary!(image_binary) do
    raise ArgumentError,
          "Image data must be a binary, got: #{inspect(image_binary)}"
  end

  @doc false
  def validate_color!(color) do
    if hex_color?(color) do
      :ok
    else
      raise ArgumentError,
            "Expected a string hex color (e.g., \"#FF0000\"), got: #{inspect(color)}"
    end
  end

  @doc """
  Validates a list of format options (see `XlsxWriter.write/5` for the
  supported options).

  Every option must be known and every value must be of the expected type
  and within range: colours must be `"#RRGGBB"`, `:font_size` a number
  greater than 0, `:rotation` an integer in `-90..90` or `270`, `:indent` an
  integer in `0..#{@max_indent}`.

  ## Raises

  - `ArgumentError` if `formats` is not a list, or an option is unknown or
    has an invalid value

  ## Examples

      iex> XlsxWriter.Validation.validate_formats!([:bold, {:bg_color, "#FF0000"}, {:font_size, 11.5}])
      :ok

      iex> XlsxWriter.Validation.validate_formats!([{:font_color, true}])
      ** (ArgumentError) Format option :font_color expects a string hex color (e.g., "#FF0000"), got: true

      iex> XlsxWriter.Validation.validate_formats!([:bolt])
      ** (ArgumentError) Unknown format option: :bolt

      iex> XlsxWriter.Validation.validate_formats!([{:rotation, 200}])
      ** (ArgumentError) Format option :rotation expects an integer between -90 and 90, or 270, got: 200

  """
  def validate_formats!(formats) when is_list(formats) do
    Enum.each(formats, &validate_format!/1)
  end

  def validate_formats!(formats) do
    raise ArgumentError,
          "Formats must be a list of format options, got: #{inspect(formats)}"
  end

  @doc """
  Validates that a value is a supported data type.

  Deprecated: `XlsxWriter.write/5` and `XlsxWriter.merge_range/7` already
  raise `ArgumentError` for unsupported values. This function still raises
  `XlsxWriter.Error`, as it did before.
  """
  @deprecated "XlsxWriter.write/5 validates values itself; this function will be removed"
  def validate_supported_type!(value) do
    case value do
      value when is_binary(value) or is_number(value) or is_boolean(value) ->
        :ok

      %Decimal{} ->
        :ok

      %Date{} ->
        :ok

      %DateTime{} ->
        :ok

      %NaiveDateTime{} ->
        :ok

      other ->
        raise XlsxWriter.Error,
              "The data type for value \"#{inspect(other)}\" is not supported."
    end
  end

  @doc """
  Validates rich string segments.

  Each segment must be a tuple of `{text, formats}` where:
  - `text` is a non-empty string
  - `formats` is a list of format options

  ## Raises

  - `ArgumentError` if segments is not a list (`"Rich string segments must
    be a list"`), is empty (`"Rich string segments cannot be empty"`), or
    contains an invalid segment

  ## Examples

      iex> XlsxWriter.Validation.validate_rich_string_segments!([{"Bold ", [:bold]}, {"Normal", []}])
      :ok

      iex> XlsxWriter.Validation.validate_rich_string_segments!([{"", [:bold]}])
      ** (ArgumentError) Rich string segment text cannot be empty

      iex> XlsxWriter.Validation.validate_rich_string_segments!([{"Bold", [:bolt]}])
      ** (ArgumentError) Unknown format option: :bolt

  """
  def validate_rich_string_segments!([]) do
    raise ArgumentError, "Rich string segments cannot be empty"
  end

  def validate_rich_string_segments!(segments) when is_list(segments) do
    Enum.each(segments, &validate_rich_string_segment!/1)
  end

  def validate_rich_string_segments!(segments) do
    raise ArgumentError,
          "Rich string segments must be a list, got: #{inspect(segments)}"
  end

  @doc false
  # Column width in Excel character units. rust_xlsxwriter caps it at 255.
  def validate_column_width!(width) do
    validate_size!(width, "Column width")
  end

  @doc false
  # Row height in points.
  def validate_row_height!(height) do
    validate_size!(height, "Row height")
  end

  @doc false
  def validate_comment_options!(opts) do
    validate_comment_author!(Keyword.get(opts, :author))
    validate_comment_visible!(Keyword.get(opts, :visible))
    validate_comment_size!(:width, Keyword.get(opts, :width))
    validate_comment_size!(:height, Keyword.get(opts, :height))
  end

  # Private helpers

  defp validate_comment_author!(nil), do: :ok

  defp validate_comment_author!(author) when is_binary(author) do
    # rust_xlsxwriter counts code points (`chars()`), not graphemes.
    if length(String.codepoints(author)) > @max_comment_author_length do
      raise ArgumentError,
            "Comment author must be at most #{@max_comment_author_length} characters, got: #{inspect(author)}"
    end

    :ok
  end

  defp validate_comment_author!(author) do
    raise ArgumentError,
          "Comment author must be a string, got: #{inspect(author)}"
  end

  defp validate_comment_visible!(visible)
       when is_boolean(visible) or is_nil(visible), do: :ok

  defp validate_comment_visible!(visible) do
    raise ArgumentError,
          "Comment :visible must be a boolean, got: #{inspect(visible)}"
  end

  defp validate_comment_size!(_option, nil), do: :ok

  # Sizes are rounded to whole pixels; rust_xlsxwriter ignores 0.
  defp validate_comment_size!(_option, size)
       when is_number(size) and round(size) >= 1 and
              round(size) <= @max_comment_size, do: :ok

  defp validate_comment_size!(option, size) do
    raise ArgumentError,
          "Comment #{inspect(option)} must be a number of pixels that rounds to at least 1, got: #{inspect(size)}"
  end

  defp validate_index!(index, _label, maximum, _unit)
       when is_integer(index) and index >= 0 and index <= maximum, do: :ok

  defp validate_index!(index, label, _maximum, _unit)
       when is_integer(index) and index < 0 do
    raise ArgumentError, "#{label} index must be non-negative, got: #{index}"
  end

  defp validate_index!(index, label, maximum, unit) when is_integer(index) do
    raise ArgumentError,
          "#{label} index must be at most #{maximum} (Excel's last #{unit}), got: #{index}"
  end

  defp validate_index!(index, label, _maximum, _unit) do
    raise ArgumentError,
          "#{label} index must be an integer, got: #{inspect(index)}"
  end

  defp validate_order!(first, last, _unit) when first <= last, do: :ok

  defp validate_order!(first, last, unit) do
    raise ArgumentError,
          "First #{unit} (#{first}) must not be after last #{unit} (#{last})"
  end

  defp validate_size!(size, _label) when is_number(size) and size >= 0, do: :ok

  defp validate_size!(size, label) do
    raise ArgumentError,
          "#{label} must be a non-negative number, got: #{inspect(size)}"
  end

  defp validate_format!(flag) when flag in @flag_formats, do: :ok

  defp validate_format!({option, color}) when option in @color_formats do
    if !hex_color?(color) do
      format_error!(
        option,
        "a string hex color (e.g., \"#FF0000\")",
        color
      )
    end
  end

  defp validate_format!({option, style}) when option in @border_formats do
    if style not in @border_styles do
      format_error!(option, "one of #{inspect(@border_styles)}", style)
    end
  end

  defp validate_format!({:align, alignment}) do
    if alignment not in @alignments do
      format_error!(:align, "one of #{inspect(@alignments)}", alignment)
    end
  end

  defp validate_format!({:valign, alignment}) do
    if alignment not in @vertical_alignments do
      format_error!(
        :valign,
        "one of #{inspect(@vertical_alignments)}",
        alignment
      )
    end
  end

  defp validate_format!({:pattern, pattern}) do
    if pattern not in @patterns do
      format_error!(:pattern, "one of #{inspect(@patterns)}", pattern)
    end
  end

  defp validate_format!({:underline, style}) do
    if style not in @underline_styles do
      format_error!(:underline, "one of #{inspect(@underline_styles)}", style)
    end
  end

  defp validate_format!({:num_format, format_string})
       when is_binary(format_string), do: :ok

  defp validate_format!({:font_name, name}) when is_binary(name), do: :ok

  defp validate_format!({:font_size, size}) when is_number(size) and size > 0,
    do: :ok

  defp validate_format!({:font_size, size}) do
    format_error!(:font_size, "a number greater than 0", size)
  end

  defp validate_format!({:rotation, angle})
       when is_integer(angle) and
              ((angle >= -90 and angle <= 90) or angle == 270), do: :ok

  defp validate_format!({:rotation, angle}) do
    format_error!(
      :rotation,
      "an integer between -90 and 90, or 270",
      angle
    )
  end

  defp validate_format!({:indent, level})
       when is_integer(level) and level >= 0 and level <= @max_indent, do: :ok

  defp validate_format!({:indent, level}) do
    format_error!(:indent, "an integer between 0 and #{@max_indent}", level)
  end

  defp validate_format!({option, value})
       when option in [:num_format, :font_name] do
    format_error!(option, "a string", value)
  end

  defp validate_format!(other) do
    raise ArgumentError, "Unknown format option: #{inspect(other)}"
  end

  defp format_error!(option, expectation, value) do
    raise ArgumentError,
          "Format option #{inspect(option)} expects #{expectation}, got: #{inspect(value)}"
  end

  defp validate_rich_string_segment!({"", _formats}) do
    raise ArgumentError, "Rich string segment text cannot be empty"
  end

  defp validate_rich_string_segment!({text, formats})
       when is_binary(text) and is_list(formats) do
    validate_formats!(formats)
  end

  defp validate_rich_string_segment!({text, _formats})
       when not is_binary(text) do
    raise ArgumentError,
          "Rich string segment text must be a string, got: #{inspect(text)}"
  end

  defp validate_rich_string_segment!({_text, formats}) do
    raise ArgumentError,
          "Rich string segment formats must be a list, got: #{inspect(formats)}"
  end

  defp validate_rich_string_segment!(other) do
    raise ArgumentError,
          "Rich string segment must be a {text, formats} tuple, got: #{inspect(other)}"
  end

  defp hex_color?(color) when is_binary(color),
    do: Regex.match?(~r/\A#?[[:xdigit:]]{6}\z/, color)

  defp hex_color?(_color), do: false
end
