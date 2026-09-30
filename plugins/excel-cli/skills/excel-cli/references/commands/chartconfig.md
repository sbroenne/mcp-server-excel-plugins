### chartconfig

Chart configuration - data source, series, type, title, axis labels, legend, and styling

**Actions:** `set-source-range`, `add-series`, `remove-series`, `set-chart-type`, `set-title`, `set-axis-title`, `get-axis-number-format`, `set-axis-number-format`, `show-legend`, `set-style`, `set-placement`, `set-data-labels`, `get-axis-scale`, `set-axis-scale`, `get-gridlines`, `set-gridlines`, `set-series-format`, `set-series-chart-type`, `get-plot-options`, `set-plot-options`, `set-area-format`, `list-trendlines`, `add-trendline`, `delete-trendline`, `set-trendline`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--chart-name` | Name of the chart (required) |
| `--source-range` | New data source range (e.g., Sheet1!A1:D10) (required for: set-source-range) (valid for: set-source-range) |
| `--series-name` | Display name for the series (required for: add-series) (valid for: add-series) |
| `--values-range` | Range containing series values (e.g., B2:B10) (required for: add-series) (valid for: add-series) |
| `--category-range` | Optional range for category labels (e.g., A2:A10) (valid for: add-series) |
| `--series-index` | 1-based index of the series to remove (required for: remove-series, set-series-format, set-series-chart-type, list-trendlines, add-trendline, delete-trendline, set-trendline) (valid for: remove-series, set-data-labels, set-series-format, set-series-chart-type, list-trendlines, add-trendline, delete-trendline, set-trendline) |
| `--chart-type` | New chart type to apply (required for: set-chart-type, set-series-chart-type) (valid for: set-chart-type, set-series-chart-type) |
| `--title` | Title text to display (required for: set-title, set-axis-title) (valid for: set-title, set-axis-title) |
| `--axis` | Which axis to set title for (Category, Value, SeriesAxis) (required for: set-axis-title, get-axis-number-format, set-axis-number-format, get-axis-scale, set-axis-scale, set-gridlines) (valid for: set-axis-title, get-axis-number-format, set-axis-number-format, get-axis-scale, set-axis-scale, set-gridlines) |
| `--number-format` | Excel number format code (e.g., "$#,##0", "0.00%") (required for: set-axis-number-format) (valid for: set-axis-number-format) |
| `--visible` | True to show legend, false to hide (required for: show-legend) (valid for: show-legend) |
| `--legend-position` | Optional position for the legend (valid for: show-legend) |
| `--style-id` | Excel chart style ID (1-48 for most chart types) (required for: set-style) (valid for: set-style) |
| `--placement` | Placement mode: 1=MoveAndSize, 2=Move, 3=FreeFloating (required for: set-placement) (valid for: set-placement) |
| `--print-object` | Whether the embedded chart prints with the worksheet (valid for: set-placement) |
| `--locked` | Whether the embedded chart is locked when the worksheet is protected (valid for: set-placement) |
| `--rounded-corners` | Whether the embedded chart uses rounded corners (valid for: set-placement) |
| `--show-value` | Show data values on labels (valid for: set-data-labels) |
| `--show-percentage` | Show percentage values. Only meaningful for pie and doughnut chart types; setting to true on other chart types has no visual effect. (valid for: set-data-labels) |
| `--show-series-name` | Show series name on labels (valid for: set-data-labels) |
| `--show-category-name` | Show category name on labels (valid for: set-data-labels) |
| `--show-bubble-size` | Show bubble size (bubble charts) (valid for: set-data-labels) |
| `--separator` | Separator string between label components (valid for: set-data-labels) |
| `--label-position` | Position of data labels relative to data points (valid for: set-data-labels) |
| `--minimum-scale` | Minimum axis value (null for auto) (valid for: set-axis-scale) |
| `--maximum-scale` | Maximum axis value (null for auto) (valid for: set-axis-scale) |
| `--major-unit` | Major gridline interval (null for auto) (valid for: set-axis-scale) |
| `--minor-unit` | Minor gridline interval (null for auto) (valid for: set-axis-scale) |
| `--show-major` | Show major gridlines (null to keep current) (valid for: set-gridlines) |
| `--show-minor` | Show minor gridlines (null to keep current) (valid for: set-gridlines) |
| `--marker-style` | Marker shape style (valid for: set-series-format) |
| `--marker-size` | Marker size in points (2-72) (valid for: set-series-format) |
| `--marker-background-color` | Marker fill color (#RRGGBB) (valid for: set-series-format) |
| `--marker-foreground-color` | Marker border color (#RRGGBB) (valid for: set-series-format) |
| `--invert-if-negative` | Invert colors for negative values (valid for: set-series-format) |
| `--fill-color` | Series fill color as #RRGGBB (valid for: set-series-format, set-area-format) |
| `--fill-transparency` | Series fill transparency from 0 (opaque) to 1 (transparent) (valid for: set-series-format, set-area-format) |
| `--line-color` | Series line color as #RRGGBB (valid for: set-series-format, set-area-format) |
| `--line-weight` | Series line weight in points (valid for: set-series-format, set-area-format) |
| `--plot-by` | Interpret source rows or columns as data series (valid for: set-plot-options) |
| `--display-blanks-as` | How blank cells appear: gaps, zeroes, or interpolation (valid for: set-plot-options) |
| `--plot-visible-only` | True to omit hidden rows and columns (valid for: set-plot-options) |
| `--area` | Chart area or plot area (required for: set-area-format) (valid for: set-area-format) |
| `--trendline-type` | Type of trendline (Linear, Exponential, etc.) (required for: add-trendline) (valid for: add-trendline) |
| `--order` | Polynomial order (2-6, for Polynomial type) (valid for: add-trendline) |
| `--period` | Moving average period (for MovingAverage type) (valid for: add-trendline) |
| `--forward` | Periods to extend forward (valid for: add-trendline, set-trendline) |
| `--backward` | Periods to extend backward (valid for: add-trendline, set-trendline) |
| `--intercept` | Force trendline through specific Y-intercept (valid for: add-trendline, set-trendline) |
| `--display-equation` | Display trendline equation on chart (valid for: add-trendline, set-trendline) |
| `--display-r-squared` | Display R-squared value on chart (valid for: add-trendline, set-trendline) |
| `--name` | Custom name for the trendline (valid for: add-trendline, set-trendline) |
| `--trendline-index` | 1-based index of the trendline to delete (required for: delete-trendline, set-trendline) (valid for: delete-trendline, set-trendline) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
