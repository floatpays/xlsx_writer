defmodule XlsxWriter.NoteOptions do
  @moduledoc false

  # This struct is used by the Rust NIF to configure Note objects
  defstruct author: nil, visible: nil, width: nil, height: nil
end
