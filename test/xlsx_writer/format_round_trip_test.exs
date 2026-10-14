defmodule XlsxWriter.FormatRoundTripTest do
  # Every format option that XlsxWriter.Validation accepts must also be
  # decoded by the NIF and end up in styles.xml. This keeps the Elixir lists
  # of allowed values and the Rust enums in sync.
  use ExUnit.Case, async: true

  import XlsxWriter.TestHelpers

  alias XlsxWriter.Validation

  @color "#123456"
  @border_color_options [
    :border_color,
    :border_top_color,
    :border_bottom_color,
    :border_left_color,
    :border_right_color
  ]

  test "every accepted format option is written by the NIF" do
    for format <- Validation.format_examples() do
      formats = with_border_style(format)
      assert Validation.validate_formats!(formats) == :ok

      sheet =
        XlsxWriter.write(XlsxWriter.new_sheet("S"), 0, 0, "x", format: formats)

      assert {:ok, content} = XlsxWriter.generate([sheet]),
             "the NIF rejected #{inspect(format)}"

      case expected_style(format) do
        nil ->
          :ok

        expected ->
          assert xlsx_part(content, "xl/styles.xml") =~ expected,
                 "#{inspect(format)} is missing #{expected} in styles.xml"
      end
    end
  end

  # Border colours are only written for a border with a style.
  defp with_border_style({option, _color} = format)
       when option in @border_color_options, do: [{:border, :thin}, format]

  defp with_border_style(format), do: [format]

  defp expected_style(:bold), do: "<b/>"
  defp expected_style(:italic), do: "<i/>"
  defp expected_style(:strikethrough), do: "<strike/>"
  defp expected_style(:superscript), do: ~s(<vertAlign val="superscript"/>)
  defp expected_style(:subscript), do: ~s(<vertAlign val="subscript"/>)
  defp expected_style(:text_wrap), do: ~s(wrapText="1")
  defp expected_style(:shrink), do: ~s(shrinkToFit="1")
  defp expected_style({:bg_color, @color}), do: ~s(<fgColor rgb="FF123456"/>)

  defp expected_style({option, @color})
       when option == :font_color or option in @border_color_options,
       do: ~s(<color rgb="FF123456"/>)

  defp expected_style({:border, style}), do: border_xml("left", style)
  defp expected_style({:border_top, style}), do: border_xml("top", style)
  defp expected_style({:border_bottom, style}), do: border_xml("bottom", style)
  defp expected_style({:border_left, style}), do: border_xml("left", style)
  defp expected_style({:border_right, style}), do: border_xml("right", style)
  defp expected_style({:align, alignment}), do: ~s(horizontal="#{alignment}")
  # Bottom is Excel's default vertical alignment, so nothing is written.
  defp expected_style({:valign, :bottom}), do: nil
  defp expected_style({:valign, alignment}), do: ~s(vertical="#{alignment}")
  # No pattern is the default fill.
  defp expected_style({:pattern, :none}), do: nil
  defp expected_style({:pattern, pattern}), do: ~s(patternType="#{pattern}")
  defp expected_style({:underline, :single}), do: "<u/>"

  defp expected_style({:underline, style}),
    do: ~s(<u val="#{lower_camel_case(style)}"/>)

  defp expected_style({:num_format, format}), do: ~s(formatCode="#{format}")
  defp expected_style({:font_name, name}), do: ~s(<name val="#{name}"/>)
  defp expected_style({:font_size, size}), do: ~s(<sz val="#{size}"/>)
  # Excel stores -90..-1 as 91..180 and stacked text as 255.
  defp expected_style({:rotation, -90}), do: ~s(textRotation="180")
  defp expected_style({:rotation, 270}), do: ~s(textRotation="255")
  defp expected_style({:rotation, angle}), do: ~s(textRotation="#{angle}")
  # An indent of 0 is the default.
  defp expected_style({:indent, 0}), do: nil
  defp expected_style({:indent, level}), do: ~s(indent="#{level}")

  defp border_xml(side, style),
    do: ~s(<#{side} style="#{lower_camel_case(style)}">)

  defp lower_camel_case(atom) do
    <<first::utf8, rest::binary>> = atom |> Atom.to_string() |> Macro.camelize()
    String.downcase(<<first::utf8>>) <> rest
  end
end
