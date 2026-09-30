### tablecolumn

Table column, filtering, and sorting operations for Excel Tables (ListObjects)

**Actions:** `apply-filter`, `apply-filter-values`, `clear-filters`, `get-filters`, `add-column`, `remove-column`, `rename-column`, `get-structured-reference`, `sort`, `sort-multi`, `get-column-number-format`, `set-column-number-format`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--table-name` | Name of the Excel table (required) |
| `--column-name` | Name of the column to filter (required for: apply-filter, apply-filter-values, add-column, remove-column, sort, get-column-number-format, set-column-number-format) (valid for: apply-filter, apply-filter-values, add-column, remove-column, get-structured-reference, sort, get-column-number-format, set-column-number-format) |
| `--criteria` | Filter criteria string (e.g., '>100', '=Active', '<>Closed') (required for: apply-filter) (valid for: apply-filter) |
| `--values` | List of exact values to include in the filter (required for: apply-filter-values) (valid for: apply-filter-values) (JSON format) |
| `--position` | 1-based column position (optional, defaults to end of table) (valid for: add-column) |
| `--old-name` | Current column name (required for: rename-column) (valid for: rename-column) |
| `--new-name` | New column name (required for: rename-column) (valid for: rename-column) |
| `--region` | Table region: 'Data', 'Headers', 'Totals', or 'All' (required for: get-structured-reference) (valid for: get-structured-reference) |
| `--ascending` | Sort order: true = ascending (A-Z, 0-9), false = descending (default: true) (valid for: sort) |
| `--sort-columns` | List of sort specifications: [{columnName: 'Col1', ascending: true}, ...] - applied in order (required for: sort-multi) (valid for: sort-multi) (JSON format) |
| `--format-code` | Number format code in US locale (e.g., '#,##0.00', '0%', 'yyyy-mm-dd') (required for: set-column-number-format) (valid for: set-column-number-format) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
