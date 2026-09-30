### window

Control Excel window visibility, position, state, status bar, and worksheet-specific views

**Actions:** `show`, `hide`, `bring-to-front`, `get-info`, `set-state`, `set-position`, `arrange`, `set-status-bar`, `clear-status-bar`, `get-view`, `freeze-panes`, `unfreeze-panes`, `set-split`, `set-zoom`, `set-display-options`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--window-state` | Window state: 'normal', 'minimized', or 'maximized' (required for: set-state) (valid for: set-state) |
| `--left` | Window left position in points (valid for: set-position) |
| `--top` | Window top position in points (valid for: set-position) |
| `--width` | Window width in points (valid for: set-position) |
| `--height` | Window height in points (valid for: set-position) |
| `--preset` | Preset name: 'left-half', 'right-half', 'top-half', 'bottom-half', 'center', 'full-screen' (required for: arrange) (valid for: arrange) |
| `--text` | Status bar text to display (e.g. "Building PivotTable from Sales data...") (required for: set-status-bar) (valid for: set-status-bar) |
| `--sheet` | Worksheet whose view should be inspected (required for: get-view, freeze-panes, unfreeze-panes, set-split, set-zoom, set-display-options) (valid for: get-view, freeze-panes, unfreeze-panes, set-split, set-zoom, set-display-options) |
| `--frozen-rows` | Number of rows to freeze from the top (0-1,048,575) (valid for: freeze-panes) |
| `--frozen-columns` | Number of columns to freeze from the left (0-16,383) (valid for: freeze-panes) |
| `--split-rows` | Number of rows above the horizontal split (0-1,048,575) (valid for: set-split) |
| `--split-columns` | Number of columns left of the vertical split (0-16,383) (valid for: set-split) |
| `--zoom` | Zoom percentage from 10 through 400 (required for: set-zoom) (valid for: set-zoom) |
| `--show-gridlines` | Whether to display cell gridlines (valid for: set-display-options) |
| `--show-headings` | Whether to display row and column headings (valid for: set-display-options) |
| `--show-outline-symbols` | Whether to display outline level symbols (valid for: set-display-options) |
| `--show-formulas` | Whether to display formulas instead of their calculated values (valid for: set-display-options) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
