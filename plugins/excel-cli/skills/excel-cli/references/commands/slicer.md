### slicer

Slicer visual filters for PivotTables and Excel Tables

**Actions:** `create-slicer`, `list-slicers`, `set-slicer-selection`, `delete-slicer`, `create-table-slicer`, `list-table-slicers`, `set-table-slicer-selection`, `delete-table-slicer`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--pivot-table-name` | Name of the PivotTable to create slicer for (required for: create-slicer) (valid for: create-slicer, list-slicers) |
| `--field-name` | Name of the field to use for the slicer (required for: create-slicer) (valid for: create-slicer) |
| `--slicer-name` | Name for the new slicer (required for: create-slicer, set-slicer-selection, delete-slicer, create-table-slicer, set-table-slicer-selection, delete-table-slicer) (valid for: create-slicer, set-slicer-selection, delete-slicer, create-table-slicer, set-table-slicer-selection, delete-table-slicer) |
| `--destination-sheet` | Worksheet where slicer will be placed (required for: create-slicer, create-table-slicer) (valid for: create-slicer, create-table-slicer) |
| `--position` | Top-left cell position for the slicer (e.g., "H2") (required for: create-slicer, create-table-slicer) (valid for: create-slicer, create-table-slicer) |
| `--selected-items` | Items to select (show in PivotTable) (required for: set-slicer-selection, set-table-slicer-selection) (valid for: set-slicer-selection, set-table-slicer-selection) (JSON format) |
| `--clear-first` | If true, clears existing selection before setting new items (default: true) (valid for: set-slicer-selection, set-table-slicer-selection) |
| `--table-name` | Name of the Excel Table (required for: create-table-slicer) (valid for: create-table-slicer, list-table-slicers) |
| `--column-name` | Name of the column to use for the slicer (required for: create-table-slicer) (valid for: create-table-slicer) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
