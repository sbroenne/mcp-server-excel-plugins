### diag

Diagnostic commands for testing CLI/MCP infrastructure without Excel

**Actions:** `ping`, `echo`, `validate-params`

| Parameter | Description |
|-----------|-------------|
| `--message` | The message to echo back (required) (required for: echo) (valid for: echo) |
| `--tag` | Optional tag to include in the response (valid for: echo) |
| `--name` | Required name parameter (required for: validate-params) (valid for: validate-params) |
| `--count` | Required integer parameter (required for: validate-params) (valid for: validate-params) |
| `--label` | Optional label parameter (valid for: validate-params) |
| `--verbose` | Optional boolean flag (default: false) (valid for: validate-params) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
