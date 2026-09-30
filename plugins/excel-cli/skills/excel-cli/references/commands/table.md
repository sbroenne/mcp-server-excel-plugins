### table

Excel Tables (ListObjects) - lifecycle and data operations

**Actions:** `list`, `preflight`, `create`, `rename`, `delete`, `read`, `resize`, `toggle-totals`, `set-column-total`, `append`, `get-data`, `set-style`, `add-to-data-model`, `create-from-dax`, `update-dax`, `get-dax`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet containing the proposed table (required for: preflight, create, create-from-dax) (valid for: preflight, create, create-from-dax) |
| `--table-name` | Name for the proposed table (must be unique in workbook) (required for: preflight, create, rename, delete, read, resize, toggle-totals, set-column-total, append, get-data, set-style, add-to-data-model, create-from-dax, update-dax, get-dax) (valid for: preflight, create, rename, delete, read, resize, toggle-totals, set-column-total, append, get-data, set-style, add-to-data-model, create-from-dax, update-dax, get-dax) |
| `--range` | Cell range address, or one cell to expand to its CurrentRegion (required for: preflight, create) (valid for: preflight, create) |
| `--has-headers` | True if the first row contains column headers (default: true) (valid for: preflight, create) |
| `--table-style` | Table style name (e.g., 'TableStyleMedium2', 'TableStyleLight1'). Optional. (required for: set-style) (valid for: create, set-style) |
| `--new-name` | New name for the table (must be unique in workbook) (required for: rename) (valid for: rename) |
| `--new-range` | New range address (e.g., 'A1:F20') (required for: resize) (valid for: resize) |
| `--show-totals` | True to show totals row, false to hide (required for: toggle-totals) (valid for: toggle-totals) |
| `--column-name` | Name of the column to set total function on (required for: set-column-total) (valid for: set-column-total) |
| `--total-function` | Totals function name: Sum, Count, Average, Min, Max, CountNums, StdDev, Var, None (required for: set-column-total) (valid for: set-column-total) |
| `--rows` | 2D array of row data to append - column order must match table columns. Optional if rowsFile is provided. (valid for: append) (JSON format) |
| `--rows-file` | Path to a JSON or CSV file containing the rows to append. JSON: 2D array. CSV: rows/columns. Alternative to inline rows parameter. (valid for: append) |
| `--visible-only` | True to return only visible (non-filtered) rows; false for all rows (default: false) (valid for: get-data) |
| `--strip-bracket-column-names` | When true, renames source table columns that contain literal bracket characters (removes brackets) beforeadding to the Data Model. This modifies the Excel table column headers in the worksheet. (valid for: add-to-data-model) |
| `--dax-query` | DAX EVALUATE query (e.g., 'EVALUATE Sales' or 'EVALUATE SUMMARIZE(...) ') (required for: create-from-dax, update-dax) (valid for: create-from-dax, update-dax) |
| `--target-cell` | Target cell address for table placement (default: 'A1') (valid for: create-from-dax) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
