### sheet

Worksheet lifecycle management: create, rename, copy, delete, move, list sheets

**Actions:** `list`, `create`, `rename`, `copy`, `delete`, `move`, `copy-to-file`, `move-to-file`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--file-path` | Optional file path when batch contains multiple workbooks. If omitted, uses primary workbook. (valid for: list, create) |
| `--sheet` | Name for the new worksheet (required for: create, delete, move) (valid for: create, delete, move) |
| `--old-name` | Current name of the worksheet (required for: rename) (valid for: rename) |
| `--new-name` | New name for the worksheet (required for: rename) (valid for: rename) |
| `--source-name` | Name of the source worksheet (required for: copy) (valid for: copy) |
| `--target-name` | Name for the copied worksheet (required for: copy) (valid for: copy) |
| `--before-sheet` | Optional: Name of sheet to position before (valid for: move, copy-to-file, move-to-file) |
| `--after-sheet` | Optional: Name of sheet to position after (valid for: move, copy-to-file, move-to-file) |
| `--source-file` | Full path to the source workbook (required for: copy-to-file, move-to-file) (valid for: copy-to-file, move-to-file) |
| `--source-sheet` | Name of the sheet to copy (required for: copy-to-file, move-to-file) (valid for: copy-to-file, move-to-file) |
| `--target-file` | Full path to the target workbook (required for: copy-to-file, move-to-file) (valid for: copy-to-file, move-to-file) |
| `--target-sheet-name` | Optional: New name for the copied sheet (default: keeps original name) (valid for: copy-to-file) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
