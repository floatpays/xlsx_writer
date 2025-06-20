%{
  configs: [
    %{
      name: "default",
      files: %{
        included: ["lib/", "test/", "config/"],
        excluded: []
      },
      strict: true,
      checks: []
    }
  ]
}
