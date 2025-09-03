defmodule XlsxWriter.MixProject do
  use Mix.Project

  @github_url "https://github.com/floatpays/xlsx_writer"
  @version "0.5.0"

  def project do
    [
      app: :xlsx_writer,
      version: @version,
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
      {:decimal, "~> 2.0"},
      {:rustler_precompiled, "~> 0.8"},
      {:rustler, "~> 0.36.1", runtime: false},

      # Dev tools
      {:credo, "~> 1.4", only: [:dev], runtime: false},
      {:quokka, "~> 2.6", only: [:dev], runtime: false},
      {:dialyxir, "~> 1.3", only: [:dev, :test], runtime: false},
      {:ex_doc, ">= 0.0.0", only: :dev, runtime: false},
      {:igniter, "~> 0.5", only: [:dev]}
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
        "native/xlsx_writer/.cargo",
        "native/xlsx_writer/src",
        "native/xlsx_writer/Cargo*",
        "checksum-*.exs"
      ],
      maintainers: ["Wilhelm H Kirschbaum", "Willem Odendaal"],
      licenses: ["MIT"],
      links: %{"GitHub" => @github_url}
    ]
  end

  defp description do
    "A fast Elixir library for writing Excel (.xlsx) files using Rust. Built with the rust_xlsxwriter crate via Rustler NIF for high performance spreadsheet generation."
  end
end
