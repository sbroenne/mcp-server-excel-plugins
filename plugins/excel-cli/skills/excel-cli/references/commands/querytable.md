### querytable

Worksheet QueryTable lifecycle and configuration for local COM text, CSV, and legacy web imports

**Actions:** `list`, `view`, `create-text`, `create-web`, `set-properties`, `refresh`, `get-refresh-status`, `cancel-refresh`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Worksheet containing the QueryTable (required for: view, create-text, create-web, set-properties, refresh, get-refresh-status, cancel-refresh, delete) (valid for: view, create-text, create-web, set-properties, refresh, get-refresh-status, cancel-refresh, delete) |
| `--query-table-name` | Name of the QueryTable (required for: view, create-text, create-web, set-properties, refresh, get-refresh-status, cancel-refresh, delete) (valid for: view, create-text, create-web, set-properties, refresh, get-refresh-status, cancel-refresh, delete) |
| `--source-path` | Full path to a readable local text or CSV file (required for: create-text) (valid for: create-text) |
| `--destination-address` | Top-left cell for the imported data (required for: create-text, create-web) (valid for: create-text, create-web) |
| `--delimiter` | Single-character field separator; defaults to comma (valid for: create-text) |
| `--text-qualifier` | Text quoting: double-quote, single-quote, or none (valid for: create-text) |
| `--encoding` | Windows code page; 65001 is UTF-8 (valid for: create-text) |
| `--has-headers` | Whether the first row contains column headings (valid for: create-text) |
| `--url` | URL of the legacy HTML web source (required for: create-web) (valid for: create-web) |
| `--selection-type` | Web selection: entire-page, all-tables, or specified-tables (valid for: create-web) |
| `--web-tables` | Comma-separated table names or indices for specified-tables selection (valid for: create-web) |
| `--formatting` | Imported web formatting: none, rich-text, or all (valid for: create-web) |
| `--background-query` | Enable or disable background refresh (valid for: set-properties) |
| `--refresh-on-file-open` | Refresh automatically when the workbook opens (valid for: set-properties) |
| `--refresh-period` | Automatic refresh interval in minutes; zero disables timed refresh (valid for: set-properties) |
| `--adjust-column-width` | Resize columns to fit refreshed data (valid for: set-properties) |
| `--preserve-formatting` | Preserve cell formatting when refreshing (valid for: set-properties) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
