### vba

VBA module and procedure operations for macro-enabled workbooks (.xlsm)

**Actions:** `list`, `view`, `import`, `update`, `run`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--module-name` | Name of the VBA module (required for: view, import, update, delete) (valid for: view, import, update, delete) |
| `--vba-code` | VBA code. Public callers must supply either inline vbaCode or a readable vbaCodeFile, not both. (required for: import, update) (valid for: import, update) |
| `--vba-code-file` | Path to a readable file containing vbaCode; use instead of inline vbaCode, not together (valid for: import, update) |
| `--procedure-name` | Name of the procedure to run (for example "Module1.MySub") (required for: run) (valid for: run) |
| `--timeout` | Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch (valid for: run) |
| `--parameters` | Optional parameters to pass to the procedure (valid for: run) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
