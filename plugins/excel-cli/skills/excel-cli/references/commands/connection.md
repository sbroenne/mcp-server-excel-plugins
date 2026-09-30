### connection

Data connections (OLEDB, ODBC, ODC import)

**Actions:** `list`, `view`, `create`, `refresh`, `get-refresh-status`, `cancel-refresh`, `delete`, `load-to`, `get-properties`, `set-properties`, `test`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--connection-name` | Name of the connection to view (required for: view, create, refresh, get-refresh-status, cancel-refresh, delete, load-to, get-properties, set-properties, test) (valid for: view, create, refresh, get-refresh-status, cancel-refresh, delete, load-to, get-properties, set-properties, test) |
| `--connection-string` | OLEDB or ODBC connection string (required for: create) (valid for: create, set-properties) |
| `--command-text` | SQL query or table name (valid for: create, set-properties) |
| `--description` | Optional description for the connection (valid for: create, set-properties) |
| `--timeout` | Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch (valid for: refresh) |
| `--sheet` | Target worksheet name (required for: load-to) (valid for: load-to) |
| `--background-query` | Run query in background (null to keep current) (valid for: set-properties) |
| `--refresh-on-file-open` | Refresh when file opens (null to keep current) (valid for: set-properties) |
| `--save-password` | Save password in connection (null to keep current) (valid for: set-properties) |
| `--refresh-period` | Auto-refresh interval in minutes (null to keep current) (valid for: set-properties) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
