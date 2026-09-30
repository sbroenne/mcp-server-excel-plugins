### chart

Chart lifecycle - create, read, move, and delete embedded charts

**Actions:** `list`, `read`, `create-from-range`, `create-from-table`, `create-from-pivottable`, `delete`, `move`, `fit-to-range`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--chart-name` | Name of the chart (or shape name) (required for: read, delete, move, fit-to-range) (valid for: read, create-from-range, create-from-table, create-from-pivottable, delete, move, fit-to-range) |
| `--sheet` | Target worksheet name (required for: create-from-range, create-from-table, create-from-pivottable, fit-to-range) (valid for: create-from-range, create-from-table, create-from-pivottable, fit-to-range) |
| `--source-range-address` | Data range for the chart (e.g., A1:D10) (required for: create-from-range) (valid for: create-from-range) |
| `--chart-type` | Type of chart to create (required for: create-from-range, create-from-table, create-from-pivottable) (valid for: create-from-range, create-from-table, create-from-pivottable) |
| `--left` | Left position in points from worksheet edge (valid for: create-from-range, create-from-table, create-from-pivottable, move) |
| `--top` | Top position in points from worksheet edge (valid for: create-from-range, create-from-table, create-from-pivottable, move) |
| `--width` | Chart width in points (valid for: create-from-range, create-from-table, create-from-pivottable, move) |
| `--height` | Chart height in points (valid for: create-from-range, create-from-table, create-from-pivottable, move) |
| `--target-range` | Cell range to position chart within (e.g., 'F2:K15'). PREFERRED over left/top. When set, left/top are ignored. (valid for: create-from-range, create-from-table, create-from-pivottable) |
| `--table-name` | Name of the Excel Table (required for: create-from-table) (valid for: create-from-table) |
| `--pivot-table-name` | Name of the source PivotTable (required for: create-from-pivottable) (valid for: create-from-pivottable) |
| `--range` | Range to fit the chart to (e.g., A1:D10) (required for: fit-to-range) (valid for: fit-to-range) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
