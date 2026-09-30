### workbook

Manage workbook metadata, document properties, Save As/copy operations, fixed-format exports, and external Excel links

**Actions:** `get-info`, `list-document-properties`, `get-document-property`, `set-document-property`, `delete-document-property`, `save-as`, `save-copy-as`, `export-fixed-format`, `list-external-links`, `update-external-link`, `break-external-link`, `set-protection`, `get-protection`, `set-view-options`, `get-view-options`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--include-built-in` | Include built-in document properties (valid for: list-document-properties) |
| `--include-custom` | Include custom document properties (valid for: list-document-properties) |
| `--property-name` | Document property name (required for: get-document-property, set-document-property, delete-document-property) (valid for: get-document-property, set-document-property, delete-document-property) |
| `--scope` | Property collection: built-in or custom (valid for: get-document-property, set-document-property) |
| `--value` | String value to store (required for: set-document-property) (valid for: set-document-property) |
| `--target-path` | Absolute output path in an existing directory (required for: save-as, save-copy-as, export-fixed-format) (valid for: save-as, save-copy-as, export-fixed-format) |
| `--format` | Output format: auto, xlsx, xlsm, xlsb, or xls (valid for: save-as) |
| `--overwrite` | Whether an existing output file may be replaced (valid for: save-as, save-copy-as, export-fixed-format) |
| `--format-type` | Fixed-forma t output: Pdf or Xps (valid for: export-fixed-format) |
| `--quality` | Export quality: Standard or Minimum (valid for: export-fixed-format) |
| `--include-document-properties` | Include document metadata in the exported file (valid for: export-fixed-format) |
| `--ignore-print-areas` | Export without restricting output to configured print areas (valid for: export-fixed-format) |
| `--from-page` | First page to export, 1-based; omit to start at the beginning (valid for: export-fixed-format) |
| `--to-page` | Last page to export, inclusive; omit to export through the end (valid for: export-fixed-format) |
| `--open-after-publish` | Open the exported file in its associated viewer (valid for: export-fixed-format) |
| `--link-source` | Exact source identifier returned by list-external-links (required for: update-external-link, break-external-link) (valid for: update-external-link, break-external-link) |
| `--is-protected` | True to protect workbook structure, false to unprotect it (required for: set-protection) (valid for: set-protection) |
| `--password` | Optional protection password; required to unprotect password-pr otected structure (valid for: set-protection) |
| `--display-gridlines` | Show or hide gridlines; omit to leave unchanged (valid for: set-view-options) |
| `--display-headings` | Show or hide row/column headings; omit to leave unchanged (valid for: set-view-options) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
