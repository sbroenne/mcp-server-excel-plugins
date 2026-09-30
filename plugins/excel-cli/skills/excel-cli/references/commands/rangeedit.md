### rangeedit

Range editing operations: insert/delete cells, rows, and columns; find/replace text; sort data

**Actions:** `insert-cells`, `delete-cells`, `insert-rows`, `delete-rows`, `insert-columns`, `delete-columns`, `find`, `replace`, `sort`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet containing the range (required) |
| `--range` | Cell range address where cells will be inserted (e.g., 'A1:D10') (required) |
| `--insert-shift` | Direction to shift existing cells: 'Down' or 'Right' (required for: insert-cells) (valid for: insert-cells) |
| `--delete-shift` | Direction to shift remaining cells: 'Up' or 'Left' (required for: delete-cells) (valid for: delete-cells) |
| `--search-value` | Text or value to search for (required for: find) (valid for: find) |
| `--find-options` | Search options: matchCase (default: false), matchEntireCell (default: false), searchFormulas (default: true) (required for: find) (valid for: find) |
| `--find-value` | Text or value to search for (required for: replace) (valid for: replace) |
| `--replace-value` | Text or value to replace matches with (required for: replace) (valid for: replace) |
| `--replace-options` | Replace options: matchCase (default: false), matchEntireCell (default: false), replaceAll (default: true) (required for: replace) (valid for: replace) |
| `--sort-columns` | Array of sort specifications: [{columnIndex: 1, ascending: true}, ...] - columnIndex is 1-based relative to range (required for: sort) (valid for: sort) (JSON format) |
| `--has-headers` | Whether the range has a header row to exclude from sorting (default: true) (valid for: sort) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
