defmodule XlsxWriter.Validation do
  @moduledoc """
  Validation functions for XlsxWriter inputs.

  This module provides validation helpers to ensure data integrity
  and provide helpful error messages before data reaches the Rust NIF.
  """

  @doc """
  Validates that row and column indices are non-negative.

  ## Parameters
  - `row` - The row index (0-based)
  - `col` - The column index (0-based)

  ## Raises
  - `ArgumentError` if row or column is negative

  ## Examples

      iex> XlsxWriter.Validation.validate_cell_position!(0, 0)
      :ok

      iex> XlsxWriter.Validation.validate_cell_position!(-1, 0)
      ** (ArgumentError) Row index must be non-negative, got: -1

  """
  def validate_cell_position!(row, col) do
    if row < 0 do
      raise ArgumentError, "Row index must be non-negative, got: #{row}"
    end

    if col < 0 do
      raise ArgumentError, "Column index must be non-negative, got: #{col}"
    end

    :ok
  end

  @doc """
  Validates that image binary data is not empty.

  ## Parameters
  - `image_binary` - Binary data for the image

  ## Raises
  - `ArgumentError` if binary is empty

  ## Examples

      iex> XlsxWriter.Validation.validate_image_binary!(<<1, 2, 3>>)
      :ok

      iex> XlsxWriter.Validation.validate_image_binary!(<<>>)
      ** (ArgumentError) Image binary cannot be empty

  """
  def validate_image_binary!(image_binary) do
    if byte_size(image_binary) == 0 do
      raise ArgumentError, "Image binary cannot be empty"
    end

    :ok
  end

  @doc """
  Validates format options list, ensuring color values are strings.

  Checks all color-related format options to ensure they receive
  string hex color values (e.g., "#FF0000") and not other types
  like booleans or integers.

  ## Parameters
  - `formats` - List of format tuples

  ## Raises
  - `XlsxWriter.Error` if any color option has a non-string value

  ## Examples

      iex> XlsxWriter.Validation.validate_formats!([:bold, {:bg_color, "#FF0000"}])
      :ok

      iex> XlsxWriter.Validation.validate_formats!([{:font_color, true}])
      ** (XlsxWriter.Error) Format option :font_color expects a string hex color (e.g., "#FF0000"), got: true

  """
  def validate_formats!(formats) when is_list(formats) do
    Enum.each(formats, fn
      {:bg_color, color} ->
        validate_color_string!(color, :bg_color)

      {:font_color, color} ->
        validate_color_string!(color, :font_color)

      {:border_color, color} ->
        validate_color_string!(color, :border_color)

      {:border_top_color, color} ->
        validate_color_string!(color, :border_top_color)

      {:border_bottom_color, color} ->
        validate_color_string!(color, :border_bottom_color)

      {:border_left_color, color} ->
        validate_color_string!(color, :border_left_color)

      {:border_right_color, color} ->
        validate_color_string!(color, :border_right_color)

      _ ->
        :ok
    end)

    :ok
  end

  @doc """
  Validates that a value is a supported data type.

  ## Parameters
  - `value` - The value to validate

  ## Raises
  - `XlsxWriter.Error` if the data type is not supported

  ## Examples

      iex> XlsxWriter.Validation.validate_supported_type!("string")
      :ok

      iex> XlsxWriter.Validation.validate_supported_type!(123)
      :ok

      iex> XlsxWriter.Validation.validate_supported_type!(self())
      ** (XlsxWriter.Error) The data type for value "#PID<0.123.0>" is not supported.

  """
  def validate_supported_type!(value) do
    case value do
      val when is_binary(val) ->
        :ok

      val when is_number(val) ->
        :ok

      val when is_boolean(val) ->
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

  # Private helpers

  defp validate_color_string!(value, _field) when is_binary(value), do: :ok

  defp validate_color_string!(value, field) do
    raise XlsxWriter.Error,
          "Format option #{inspect(field)} expects a string hex color (e.g., \"#FF0000\"), got: #{inspect(value)}"
  end
end
