### xmlmap

Manage workbook XML maps and exchange XML data without interactive dialogs

**Actions:** `list`, `add`, `map-range`, `import-xml`, `export-xml`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--schema` | XSD schema content. Public callers must supply either inline schema or a readable schemaFile, not both. (required for: add) (valid for: add) |
| `--schema-file` | Path to a readable file containing schema; use instead of inline schema, not together (valid for: add) |
| `--root-element-name` | Optional root element when the schema has multiple roots (valid for: add) |
| `--map-name` | Optional name to assign to the created map (required for: map-range, export-xml, delete) (valid for: add, map-range, import-xml, export-xml, delete) |
| `--sheet` | Worksheet containing the target range (required for: map-range) (valid for: map-range, import-xml) |
| `--range` | Target cell or single-column range (required for: map-range) (valid for: map-range) |
| `--xpath` | XPath to map (required for: map-range) (valid for: map-range) |
| `--selection-namespace` | Optional namespace declarations used by prefixed XPath expressions (valid for: map-range) |
| `--repeating` | Whether to create a repeating XML list mapping (valid for: map-range) |
| `--xml-data` | XML data. Public callers must supply either inline xmlData or a readable xmlDataFile, not both. (required for: import-xml) (valid for: import-xml) |
| `--xml-data-file` | Path to a readable file containing xmlData; use instead of inline xmlData, not together (valid for: import-xml) |
| `--start-cell` | Top-left destination cell for automatic mapping (valid for: import-xml) |
| `--overwrite` | Whether imported XML may overwrite mapped cells (valid for: import-xml) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
