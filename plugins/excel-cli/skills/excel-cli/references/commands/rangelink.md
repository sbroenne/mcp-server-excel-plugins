### rangelink

Hyperlink, threaded comment, and cell protection operations for Excel ranges

**Actions:** `add-hyperlink`, `update-hyperlink`, `remove-hyperlink`, `list-hyperlinks`, `get-hyperlink`, `add-threaded-comment`, `list-threaded-comments`, `add-threaded-comment-reply`, `delete-threaded-comment`, `set-cell-lock`, `get-cell-lock`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet (required) |
| `--cell-address` | Single cell address (e.g., 'A1') (required for: add-hyperlink, update-hyperlink, get-hyperlink, add-threaded-comment, list-threaded-comments, add-threaded-comment-reply, delete-threaded-comment) (valid for: add-hyperlink, update-hyperlink, get-hyperlink, add-threaded-comment, list-threaded-comments, add-threaded-comment-reply, delete-threaded-comment) |
| `--url` | Optional external URL or file path. Omit for an internal workbook link. (valid for: add-hyperlink, update-hyperlink) |
| `--display-text` | Text to display in the cell (optional, defaults to URL) (valid for: add-hyperlink, update-hyperlink) |
| `--tooltip` | Tooltip text shown on hover (optional) (valid for: add-hyperlink, update-hyperlink) |
| `--sub-address` | Optional internal workbook target such as "'Sheet2'!A1" (valid for: add-hyperlink, update-hyperlink) |
| `--range` | Cell range address to remove hyperlinks from (e.g., 'A1:D10') (required for: remove-hyperlink, set-cell-lock, get-cell-lock) (valid for: remove-hyperlink, set-cell-lock, get-cell-lock) |
| `--text` | Comment or reply text; cloud mentions and assignments are not supported (required for: add-threaded-comment, add-threaded-comment-reply) (valid for: add-threaded-comment, add-threaded-comment-reply) |
| `--locked` | Lock status: true = locked (protected when sheet protection enabled), false = unlocked (editable) (required for: set-cell-lock) (valid for: set-cell-lock) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
