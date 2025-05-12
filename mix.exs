defmodule XlsxWriter.MixProject do
  use Mix.Project

  def project do
    [
      app: :xlsx_writer,
      version: "0.2.0",
      elixir: "~> 1.18",
      start_permanent: Mix.env() == :prod,
      deps: deps(),
      description: description(),
      package: package()
    ]
  end

  # Run "mix help compile.app" to learn about applications.
  def application do
    [
      extra_applications: [:logger]
    ]
  end

  defp deps do
    [
      {:rustler_precompiled, "~> 0.8"},
      {:rustler, "~> 0.36.1", runtime: false},
      {:credo, "~> 1.4", only: [:dev, :test], runtime: false},
      {:dialyxir, "~> 1.3", only: [:dev, :test], runtime: false}
    ]
  end

  defp package do
    [
      organization: "floatpays",
      files: [
        "lib",
        "mix.exs",
        "README*",
        "LICENSE*",
        "native/rustxlsxwriter/Cargo.toml",
        "native/rustxlsxwriter/src"
      ],
      maintainers: ["Wilhelm H Kirschbaum", "Willem Odendaal"],
      licenses: ["MIT"],
      links: %{"GitHub" => "https://github.com/floatpays/xlsx_writer"}
    ]
  end

  defp description do
    "Writes Xlsx spreadsheet using the rust_xlsxwriter package, via a Rustler NIF."
  end
end
