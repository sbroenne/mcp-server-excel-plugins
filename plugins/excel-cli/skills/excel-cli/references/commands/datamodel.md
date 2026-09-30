### datamodel

Data Model (Power Pivot) - DAX measures and table management

**Actions:** `list-tables`, `list-columns`, `read-table`, `read-info`, `read-connection`, `list-measures`, `read`, `delete-measure`, `delete-table`, `rename-table`, `refresh`, `create-measure`, `update-measure`, `evaluate`, `execute-dmv`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--table-name` | Name of the table to list columns from (required for: list-columns, read-table, delete-table, create-measure) (valid for: list-columns, read-table, list-measures, delete-table, refresh, create-measure) |
| `--measure-name` | Name of the measure to get (required for: read, delete-measure, create-measure, update-measure) (valid for: read, delete-measure, create-measure, update-measure) |
| `--old-name` | Current name of the table (required for: rename-table) (valid for: rename-table) |
| `--new-name` | New name for the table (required for: rename-table) (valid for: rename-table) |
| `--timeout` | Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch (valid for: refresh) |
| `--dax-formula` | DAX formula. Public callers must supply either inline daxFormula or a readable daxFormulaFile, not both. (required for: create-measure) (valid for: create-measure, update-measure) |
| `--dax-formula-file` | Path to a readable file containing daxFormula; use instead of inline daxFormula, not together (valid for: create-measure, update-measure) |
| `--format-type` | Optional format type: General, Currency, Decimal, Percentage, or WholeNumber (case-insensitive). Null or empty defaults to General on create and keeps the existing format on update. (valid for: create-measure, update-measure) |
| `--description` | Optional: Description of the measure (valid for: create-measure, update-measure) |
| `--format-dax` | Whether to send the DAX formula to the remote daxformatter.com service before saving. Defaults to false to preserve privacy. (valid for: create-measure, update-measure) |
| `--dax-query` | DAX EVALUATE query. Public callers must supply either inline daxQuery or a readable daxQueryFile, not both. (required for: evaluate) (valid for: evaluate) |
| `--dax-query-file` | Path to a readable file containing daxQuery; use instead of inline daxQuery, not together (valid for: evaluate) |
| `--dmv-query` | DMV query in SQL-like syntax. Public callers must supply either inline dmvQuery or a readable dmvQueryFile, not both. (required for: execute-dmv) (valid for: execute-dmv) |
| `--dmv-query-file` | Path to a readable file containing dmvQuery; use instead of inline dmvQuery, not together (valid for: execute-dmv) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
