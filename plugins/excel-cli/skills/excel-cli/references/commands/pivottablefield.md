### pivottablefield

PivotTable field management: add/remove/configure fields, filtering, sorting, and grouping

**Actions:** `list-fields`, `add-row-field`, `add-column-field`, `add-value-field`, `add-filter-field`, `remove-field`, `set-field-function`, `set-field-name`, `set-field-format`, `set-field-filter`, `sort-field`, `group-by-date`, `group-by-numeric`, `group-items`, `ungroup-field`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--pivot-table-name` | Name of the PivotTable (required) |
| `--field-name` | Name of the field to add (required for: add-row-field, add-column-field, add-value-field, add-filter-field, remove-field, set-field-function, set-field-name, set-field-format, set-field-filter, sort-field, group-by-date, group-by-numeric, group-items) (valid for: add-row-field, add-column-field, add-value-field, add-filter-field, remove-field, set-field-function, set-field-name, set-field-format, set-field-filter, sort-field, group-by-date, group-by-numeric, group-items) |
| `--position` | Optional position in row area (1-based) (valid for: add-row-field, add-column-field) |
| `--aggregation-function` | Aggregation function (for Regular and OLAP auto-create mode) (required for: set-field-function) (valid for: add-value-field, set-field-function) |
| `--custom-name` | Optional custom name for the field/measure (required for: set-field-name) (valid for: add-value-field, set-field-name) |
| `--number-format` | Number format string (required for: set-field-format) (valid for: set-field-format) |
| `--selected-values` | Values to show (others will be hidden) (required for: set-field-filter) (valid for: set-field-filter) (JSON format) |
| `--direction` | Sort direction (valid for: sort-field) |
| `--interval` | Grouping interval (Months, Quarters, Years) (required for: group-by-date) (valid for: group-by-date) |
| `--start` | Starting value (null = use field minimum) (valid for: group-by-numeric) |
| `--end-value` | Ending value (null = use field maximum) (valid for: group-by-numeric) |
| `--interval-size` | Size of each group (e.g., 100 for groups of 100) (required for: group-by-numeric) (valid for: group-by-numeric) |
| `--item-names` | Exact item captions to group (required for: group-items) (valid for: group-items) (JSON format) |
| `--group-name` | Caption for the new group (required for: group-items) (valid for: group-items) |
| `--grouped-field-name` | Generated grouped field name (required for: ungroup-field) (valid for: ungroup-field) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
