### pivottable

PivotTable lifecycle management: create from various sources, list, read details, refresh, and delete

**Actions:** `list`, `read`, `create-from-range`, `create-from-table`, `create-from-datamodel`, `delete`, `refresh`, `get-cache-options`, `set-cache-options`, `drill-through`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--pivot-table-name` | Name of the PivotTable (required for: read, create-from-range, create-from-table, create-from-datamodel, delete, refresh, get-cache-options, set-cache-options, drill-through) (valid for: read, create-from-range, create-from-table, create-from-datamodel, delete, refresh, get-cache-options, set-cache-options, drill-through) |
| `--source-sheet` | Source worksheet name (required for: create-from-range) (valid for: create-from-range) |
| `--source-range` | Source range address (e.g., "A1:F100") (required for: create-from-range) (valid for: create-from-range) |
| `--destination-sheet` | Destination worksheet name (required for: create-from-range, create-from-table, create-from-datamodel) (valid for: create-from-range, create-from-table, create-from-datamodel) |
| `--destination-cell` | Destination cell address (e.g., "A1") (required for: create-from-range, create-from-table, create-from-datamodel) (valid for: create-from-range, create-from-table, create-from-datamodel) |
| `--table-name` | Name of the Excel Table (required for: create-from-table, create-from-datamodel) (valid for: create-from-table, create-from-datamodel) |
| `--timeout` | Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch (valid for: refresh) |
| `--enable-refresh` | Allow the PivotCache to refresh (valid for: set-cache-options) |
| `--refresh-on-file-open` | Refresh the PivotCache when the workbook opens (valid for: set-cache-options) |
| `--missing-items-limit` | How many deleted source items Excel retains in the cache (valid for: set-cache-options) |
| `--optimize-cache` | Optimize the cache when it is constructed (valid for: set-cache-options) |
| `--save-source-data` | Save source records with the PivotTable (valid for: set-cache-options) |
| `--cell-address` | Value-cell address on the PivotTable worksheet (for example, G4) (required for: drill-through) (valid for: drill-through) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
