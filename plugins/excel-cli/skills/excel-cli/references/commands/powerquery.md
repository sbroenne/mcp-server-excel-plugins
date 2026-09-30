### powerquery

Power Query M code and data loading

**Actions:** `list`, `view`, `refresh`, `get-load-config`, `delete`, `create`, `update`, `load-to`, `refresh-all`, `rename`, `unload`, `evaluate`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--query-name` | Name of the query to view (required for: view, refresh, get-load-config, delete, create, update, load-to, unload) (valid for: view, refresh, get-load-config, delete, create, update, load-to, unload) |
| `--timeout` | Public input is whole seconds from 0 through 2147483. Omitted or 0 uses the 30-minute data-operation default. (valid for: refresh, refresh-all) |
| `--m-code` | Raw M code. Public callers must supply either inline mCode or a readable mCodeFile, not both. (required for: create, update, evaluate) (valid for: create, update, evaluate) |
| `--m-code-file` | Path to a readable file containing mCode; use instead of inline mCode, not together (valid for: create, update, evaluate) |
| `--load-destination` | Load destination mode (required for: load-to) (valid for: create, load-to) |
| `--target-sheet` | Target worksheet name (required for LoadToTable and LoadToBoth; defaults to query name when omitted) (valid for: create, load-to) |
| `--target-cell-address` | Optional target cell address for worksheet loads (e.g., "B5"). Required when loading to an existing worksheet with other data. (valid for: create, load-to) |
| `--format-m-code` | Whether to send M code to the remote powerqueryformatter.com service before saving. Defaults to false to preserve privacy. (valid for: create, update) |
| `--refresh` | Whether to refresh data after update (default: true) (valid for: update) |
| `--old-name` | Current name of the query (required for: rename) (valid for: rename) |
| `--new-name` | New name for the query (required for: rename) (valid for: rename) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
