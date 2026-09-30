### pythoninexcel

Microsoft 365 "Python in Excel" (=PY()) formulas

**Actions:** `set-formula`, `get-result`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Worksheet name (required) |
| `--range` | Target cell address (e.g. "D1") (required) |
| `--code` | Python source code (e.g. "xl('A1:A6').sum()") (required for: set-formula) (valid for: set-formula) |
| `--return-type` | 0 = Excel Value (default), 1 = Python Object (valid for: set-formula) |
| `--max-wait-seconds` | Maximum seconds to poll for the cloud result before giving up (default: 30). Must be shorter than the session operation timeout. Returns as soon as the result is ready. (valid for: get-result) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
