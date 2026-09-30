### worksheetstyle

Worksheet styling, visibility, protection, grouping, and outline operations

**Actions:** `set-tab-color`, `get-tab-color`, `clear-tab-color`, `set-protection`, `get-protection`, `set-comment`, `get-comment`, `clear-comment`, `add-image`, `get-image-count`, `add-shape`, `get-shape-count`, `set-page-setup`, `get-page-setup`, `set-visibility`, `get-visibility`, `show`, `hide`, `very-hide`, `group`, `ungroup`, `get-outline-info`, `set-outline-settings`, `show-outline-levels`, `clear-outline`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet to color (required) |
| `--red` | Red color component (0-255) (required for: set-tab-color) (valid for: set-tab-color) |
| `--green` | Green color component (0-255) (required for: set-tab-color) (valid for: set-tab-color) |
| `--blue` | Blue color component (0-255) (required for: set-tab-color) (valid for: set-tab-color) |
| `--is-protected` | Whether the worksheet should be protected (required for: set-protection) (valid for: set-protection) |
| `--password` | Optional password for protecting/unprotecting the sheet (valid for: set-protection) |
| `--cell-address` | Cell address such as A1 (required for: set-comment, get-comment, clear-comment, add-image, add-shape) (valid for: set-comment, get-comment, clear-comment, add-image, add-shape) |
| `--text` | Cell note text to set (required for: set-comment) (valid for: set-comment) |
| `--image-path` | Absolute path to the image file on disk (required for: add-image) (valid for: add-image) |
| `--orientation` | Page orientation: 'portrait' or 'landscape' (required for: set-page-setup) (valid for: set-page-setup) |
| `--fit-to-pages-wide` | Number of pages wide to fit the printout to (valid for: set-page-setup) |
| `--fit-to-pages-tall` | Number of pages tall to fit the printout to (valid for: set-page-setup) |
| `--center-horizontally` | Whether to center the printout horizontally on the page (valid for: set-page-setup) |
| `--center-vertically` | Whether to center the printout vertically on the page (valid for: set-page-setup) |
| `--visibility` | Visibility level: 'visible', 'hidden', or 'veryhidden' (required for: set-visibility) (valid for: set-visibility) |
| `--range` | Row or column range to group (required for: group, ungroup, get-outline-info) (valid for: group, ungroup, get-outline-info) |
| `--axis` | Grouping axis: Rows or Columns (required for: group, ungroup, get-outline-info) (valid for: group, ungroup, get-outline-info) |
| `--summary-row` | Summary row position: above or below (valid for: set-outline-settings) |
| `--summary-column` | Summary column position: left or right (valid for: set-outline-settings) |
| `--automatic-styles` | Whether Excel applies automatic outline styles (valid for: set-outline-settings) |
| `--row-levels` | Optional row outline level to display (valid for: show-outline-levels) |
| `--column-levels` | Optional column outline level to display (valid for: show-outline-levels) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
