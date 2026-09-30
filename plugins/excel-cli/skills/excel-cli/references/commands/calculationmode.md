### calculationmode

Control Excel recalculation (automatic vs manual)

**Actions:** `get-mode`, `set-mode`, `calculate`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--mode` | Target calculation mode (required for: set-mode) (valid for: set-mode) |
| `--scope` | Scope: Workbook, Sheet, or Range (required for: calculate) (valid for: calculate) |
| `--sheet` | Sheet name (required for Sheet/Range scope) (valid for: calculate) |
| `--range` | Range address (required for Range scope) (valid for: calculate) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
