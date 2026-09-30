### rangeformat

Range formatting operations: apply styles, set fonts/colors/borders, add data validation, merge cells, auto-fit dimensions

**Actions:** `set-style`, `get-style`, `format-range`, `format-ranges`, `validate-range`, `get-validation`, `remove-validation`, `auto-fit-columns`, `auto-fit-rows`, `merge-cells`, `unmerge-cells`, `get-merge-info`, `set-column-width`, `set-row-height`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Name of the worksheet containing the range (required) |
| `--range` | Cell range address (e.g., 'A1:D10') (required for: set-style, get-style, format-range, validate-range, get-validation, remove-validation, auto-fit-columns, auto-fit-rows, merge-cells, unmerge-cells, get-merge-info, set-column-width, set-row-height) (valid for: set-style, get-style, format-range, validate-range, get-validation, remove-validation, auto-fit-columns, auto-fit-rows, merge-cells, unmerge-cells, get-merge-info, set-column-width, set-row-height) |
| `--style-name` | Built-in or custom style name (e.g., 'Heading 1', 'Good', 'Bad', 'Currency', 'Percent'). Use 'Normal' to reset. (required for: set-style) (valid for: set-style) |
| `--font-name` | Font family name (e.g., 'Arial', 'Calibri', 'Times New Roman') (valid for: format-range, format-ranges) |
| `--font-size` | Font size in points (e.g., 10, 11, 12, 14, 16) (valid for: format-range, format-ranges) |
| `--bold` | Whether to apply bold formatting (valid for: format-range, format-ranges) |
| `--italic` | Whether to apply italic formatting (valid for: format-range, format-ranges) |
| `--underline` | Whether to apply underline formatting (valid for: format-range, format-ranges) |
| `--font-color` | Font (foreground) color as hex '#RRGGBB' (e.g., '#FF0000' for red) (valid for: format-range, format-ranges) |
| `--fill-color` | Cell fill (background) color as hex '#RRGGBB' (e.g., '#FFFF00' for yellow) (valid for: format-range, format-ranges) |
| `--border-style` | Border line style: 'continuous', 'dash', 'dot', 'dashdot', 'dashdotdot', 'double', 'slantdashdot', 'none' (valid for: format-range, format-ranges) |
| `--border-color` | Border color as hex '#RRGGBB' (valid for: format-range, format-ranges) |
| `--border-weight` | Border weight: 'hairline', 'thin', 'medium', 'thick' (valid for: format-range, format-ranges) |
| `--horizontal-alignment` | Horizontal text alignment: 'left', 'center', 'right', 'justify', 'fill' (valid for: format-range, format-ranges) |
| `--vertical-alignment` | Vertical text alignment: 'top', 'center' (or 'middle'), 'bottom', 'justify' (valid for: format-range, format-ranges) |
| `--wrap-text` | Whether to wrap text within cells (valid for: format-range, format-ranges) |
| `--orientation` | Text rotation in degrees (-90 to 90, or 255 for vertical) (valid for: format-range, format-ranges) |
| `--range-addresses` | Cell range addresses to format (e.g., 'A1:D1', 'A3:D3') (required for: format-ranges) (valid for: format-ranges) |
| `--number-format` | Excel number format code applied to all target ranges (e.g., '0.00%' for percentage, '$#,##0.00' for currency, 'm/d/yyyy' for date). LLMs know Excel format codes natively. (valid for: format-ranges) |
| `--validation-type` | Data validation type: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom' (required for: validate-range) (valid for: validate-range) |
| `--validation-operator` | Validation comparison operator: 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual' (valid for: validate-range) |
| `--formula1` | First validation formula/value - for list validation use range '=$A$1:$A$10' or inline '"A,B,C"' (valid for: validate-range) |
| `--formula2` | Second validation formula/value - required only for 'between' and 'notBetween' operators (valid for: validate-range) |
| `--show-input-message` | Whether to show input message when cell is selected (default: false) (valid for: validate-range) |
| `--input-title` | Title for the input message popup (valid for: validate-range) |
| `--input-message` | Text for the input message popup (valid for: validate-range) |
| `--show-error-alert` | Whether to show error alert on invalid input (default: true) (valid for: validate-range) |
| `--error-style` | Error alert style: 'stop' (prevents entry), 'warning' (allows override), 'information' (allows entry) (valid for: validate-range) |
| `--error-title` | Title for the error alert popup (valid for: validate-range) |
| `--error-message` | Text for the error alert popup (valid for: validate-range) |
| `--ignore-blank` | Whether to allow blank cells in validation (default: true) (valid for: validate-range) |
| `--show-dropdown` | Whether to show dropdown arrow for list validation (default: true) (valid for: validate-range) |
| `--column-width` | Width in points (1 point = 1/72 inch, approx 0.35mm). Standard width ~8.43 points. Range: 0.25-409 points. (required for: set-column-width) (valid for: set-column-width) |
| `--row-height` | Height in points (1 point = 1/72 inch, approx 0.35mm). Default row height ~15 points. Range: 0-409 points. (required for: set-row-height) (valid for: set-row-height) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
