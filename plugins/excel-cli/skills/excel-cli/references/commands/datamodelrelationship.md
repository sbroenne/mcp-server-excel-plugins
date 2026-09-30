### datamodelrelationship

Data Model relationships - link tables for cross-table DAX calculations

**Actions:** `list-relationships`, `read-relationship`, `create-relationship`, `update-relationship`, `delete-relationship`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--from-table` | Source table name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) (valid for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--from-column` | Source column name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) (valid for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--to-table` | Target table name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) (valid for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--to-column` | Target column name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) (valid for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--active` | Whether the relationship should be active (default: true) (required for: update-relationship) (valid for: create-relationship, update-relationship) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
