defmodule XlsxWriter.Error do
  @moduledoc false

  # Since 0.11.0 validation errors are `ArgumentError`. This exception is only
  # raised by the deprecated `XlsxWriter.Validation.validate_supported_type!/1`
  # and is kept so that existing `rescue XlsxWriter.Error` clauses compile.

  defexception [:message]

  def new(message) when is_binary(message) do
    %__MODULE__{message: message}
  end
end
