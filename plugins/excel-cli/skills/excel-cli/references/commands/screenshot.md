### screenshot

Capture Excel worksheet content as images for visual verification

**Actions:** `capture`, `capture-sheet`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Worksheet name (null for active sheet) |
| `--range` | Range to capture (e.g., "A1:F20") (valid for: capture) |
| `--quality` | Image quality: Medium (default, JPEG 75% scale), High (PNG full scale), Low (JPEG 50% scale) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
