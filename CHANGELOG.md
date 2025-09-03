# v0.5.0

## breaking

- Rename module `XlsxWriter.Workbook` to `XlsxWriter` - this simplifies the API by removing the nested module structure. All functions that were previously called as `XlsxWriter.Workbook.function_name()` should now be called as `XlsxWriter.function_name()`

## improvements

- Cleaner, more intuitive module structure
- Simplified imports - no need for `alias XlsxWriter.Workbook` anymore

# v0.4.0

## breaking

- Unify write and write_with_format functions - the API has been simplified with format options now passed as part of the options map

## improvements

- Add comprehensive formatting options documentation to README
- Update package description for better clarity
- Clean up and improve README documentation with advanced usage examples

# v0.3.6

- No changes, just a release fix

# v0.3.5

## improvements

- Clean up README documentation and formatting
- Update rust_xlsxwriter dependency from 0.88.0 to 0.90.0

# v0.3.0

## breaking

- XlsxWriter returns a binary string, not IO data anymore
