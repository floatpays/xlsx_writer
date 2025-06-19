defmodule XlsxWriter.Error do
  @moduledoc false

  defexception [:message]

  def new(message) when is_binary(message) do
    %__MODULE__{message: message}
  end
end
