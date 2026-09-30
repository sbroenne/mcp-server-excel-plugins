### conditionalformat

Conditional formatting - visual rules based on cell values

**Actions:** `add-rule`, `clear-rules`, `list-rules`, `list-worksheet-rules`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Sheet name (empty for active sheet) (required) |
| `--range` | Range address (A1 notation or named range) (required for: add-rule, clear-rules, list-rules) (valid for: add-rule, clear-rules, list-rules) |
| `--rule-type` | Rule type: cellValue (or cell-value), expression, colorScale, dataBar, top10, iconSet, uniqueValues, blanksCondition, timePeriod, aboveAverage. Both camelCase and kebab-case accepted. (required for: add-rule) (valid for: add-rule) |
| `--operator-type` | Required for cellValue rules. XlFormatConditionOpe rator: equal, notEqual, greater, less, greaterEqual, lessEqual, between, notBetween (valid for: add-rule) |
| `--formula1` | Required for cellValue and expression rules. First formula/value for condition (valid for: add-rule) |
| `--formula2` | Required for between/notBetween cellValue rules. Second formula/value (valid for: add-rule) |
| `--interior-color` | Fill color (#RRGGBB or color index) (valid for: add-rule) |
| `--interior-pattern` | Interior pattern (1=Solid, -4142=None, 9=Gray50, etc.) (valid for: add-rule) |
| `--font-color` | Font color (#RRGGBB or color index) (valid for: add-rule) |
| `--font-bold` | Bold font (valid for: add-rule) |
| `--font-italic` | Italic font (valid for: add-rule) |
| `--border-style` | Border style: none, continuous, dash, dot, etc. (valid for: add-rule) |
| `--border-color` | Border color (#RRGGBB or color index) (valid for: add-rule) |
| `--color-scale-min-type` | colorScale minimum stop type: minimum, number, percent, percentile, formula (valid for: add-rule) |
| `--color-scale-min-value` | colorScale minimum stop value (for number/percent/perce ntile/formula) (valid for: add-rule) |
| `--color-scale-min-color` | colorScale minimum stop color (#RRGGBB) (valid for: add-rule) |
| `--color-scale-mid-type` | colorScale midpoint stop type (supply to create a 3-color scale) (valid for: add-rule) |
| `--color-scale-mid-value` | colorScale midpoint stop value (valid for: add-rule) |
| `--color-scale-mid-color` | colorScale midpoint stop color (#RRGGBB) (valid for: add-rule) |
| `--color-scale-max-type` | colorScale maximum stop type: maximum, number, percent, percentile, formula (valid for: add-rule) |
| `--color-scale-max-value` | colorScale maximum stop value (valid for: add-rule) |
| `--color-scale-max-color` | colorScale maximum stop color (#RRGGBB) (valid for: add-rule) |
| `--data-bar-color` | dataBar fill color (#RRGGBB) (valid for: add-rule) |
| `--data-bar-negative-color` | dataBar negative-value bar color (#RRGGBB) (valid for: add-rule) |
| `--data-bar-direction` | dataBar fill direction: context, leftToRight, rightToLeft (valid for: add-rule) |
| `--data-bar-show-value` | dataBar show the cell value alongside the bar (valid for: add-rule) |
| `--data-bar-min-type` | dataBar minimum point type: automaticMinimum, minimum, number, percent, percentile, formula (valid for: add-rule) |
| `--data-bar-min-value` | dataBar minimum point value (valid for: add-rule) |
| `--data-bar-max-type` | dataBar maximum point type: automaticMaximum, maximum, number, percent, percentile, formula (valid for: add-rule) |
| `--data-bar-max-value` | dataBar maximum point value (valid for: add-rule) |
| `--icon-set-id` | iconSet id: 3Arrows, 3TrafficLights1, 4Ratings, 5Quarters, etc. (valid for: add-rule) |
| `--icon-set-reverse` | iconSet reverse icon order (valid for: add-rule) |
| `--icon-set-show-icon-only` | iconSet show only the icon (hide the value) (valid for: add-rule) |
| `--icon-threshold1-type` | iconSet threshold 1 type: percent, number, percentile, formula (valid for: add-rule) |
| `--icon-threshold1-value` | iconSet threshold 1 value (valid for: add-rule) |
| `--icon-threshold2-type` | iconSet threshold 2 type (valid for: add-rule) |
| `--icon-threshold2-value` | iconSet threshold 2 value (valid for: add-rule) |
| `--icon-threshold3-type` | iconSet threshold 3 type (valid for: add-rule) |
| `--icon-threshold3-value` | iconSet threshold 3 value (valid for: add-rule) |
| `--icon-threshold4-type` | iconSet threshold 4 type (valid for: add-rule) |
| `--icon-threshold4-value` | iconSet threshold 4 value (valid for: add-rule) |
| `--rank` | top10 rank (number of values, or percent when top10Percent is true) (valid for: add-rule) |
| `--top10-percent` | top10 treat rank as a percentage (valid for: add-rule) |
| `--top-bottom` | top10 direction: top or bottom (valid for: add-rule) |
| `--above-below` | aboveAverage selector: aboveAverage, belowAverage, aboveStdDev, belowStdDev, equalAboveAverage, equalBelowAverage (valid for: add-rule) |
| `--date-period` | timePeriod period: today, yesterday, tomorrow, last7Days, thisWeek, lastWeek, nextWeek, thisMonth, lastMonth, nextMonth (valid for: add-rule) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
