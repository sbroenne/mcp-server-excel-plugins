### range

Core range operations: get/set values and formulas, copy ranges, clear content, and discover data regions

**Actions:** `get-values`, `set-values`, `get-formulas`, `set-formulas`, `validate-formulas`, `clear-all`, `clear-contents`, `clear-formats`, `copy`, `copy-values`, `copy-formulas`, `get-number-formats`, `set-number-format`, `set-number-formats`, `get-used-range`, `get-current-region`, `get-info`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet containing the range - REQUIRED for cell addresses, use empty string for named ranges only (required for: get-values, set-values, get-formulas, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-used-range, get-current-region, get-info) (valid for: get-values, set-values, get-formulas, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-used-range, get-current-region, get-info) |
| `--range` | Cell range address (e.g., 'A1', 'A1:D10', 'B:D') or named range name (e.g., 'SalesData') (required for: get-values, set-values, get-formulas, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-info) (valid for: get-values, set-values, get-formulas, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-info) |
| `--values` | 2D array of values to set - rows are outer array, columns are inner array (e.g., [[1,2,3],[4,5,6]] for 2 rows x 3 cols). Strict ISO dates such as "2025-01-15" become native Excel dates. Optional if valuesFile is provided. (valid for: set-values) (JSON format) |
| `--values-file` | Path to a JSON or CSV file containing the values. JSON: 2D array. CSV: rows/columns. Alternative to inline values parameter. (valid for: set-values) |
| `--formulas` | 2D array of formulas to set - include '=' prefix (e.g., [['=A1+B1', '=SUM(A:A)'], ['=C1*2', '=AVERAGE(B:B)']]). Optional if formulasFile is provided. (valid for: set-formulas, validate-formulas) (JSON format) |
| `--formulas-file` | Path to a JSON file containing the formulas as a 2D array. Alternative to inline formulas parameter. (valid for: set-formulas, validate-formulas) |
| `--source-sheet` | Source worksheet name for copy operations (required for: copy, copy-values, copy-formulas) (valid for: copy, copy-values, copy-formulas) |
| `--source-range` | Source range address for copy operations (e.g., 'A1:D10') (required for: copy, copy-values, copy-formulas) (valid for: copy, copy-values, copy-formulas) |
| `--target-sheet` | Target worksheet name for copy operations (required for: copy, copy-values, copy-formulas) (valid for: copy, copy-values, copy-formulas) |
| `--target-range` | Target range address - can be single cell for paste destination (e.g., 'A1') (required for: copy, copy-values, copy-formulas) (valid for: copy, copy-values, copy-formulas) |
| `--format-code` | Number format code in US locale (e.g., '#,##0.00' for numbers, 'mm/dd/yyyy' for dates, '0.00%' for percentages, 'General' for default, '@' for text) (required for: set-number-format) (valid for: set-number-format) |
| `--formats` | 2D array of format codes - same dimensions as target range (e.g., [['#,##0.00', '0.00%'], ['mm/dd/yyyy', 'General']]). Optional if formatsFile is provided. (valid for: set-number-formats) (JSON format) |
| `--formats-file` | Path to a JSON file containing 2D array of format codes. Alternative to inline formats parameter. (valid for: set-number-formats) |
| `--cell-address` | Single cell address (e.g., 'B5') - expands to contiguous data region around this cell (required for: get-current-region) (valid for: get-current-region) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
