### batch

Execute multiple commands from a JSON file or stdin. Outputs NDJSON (one result per line)

| Parameter | Description |
|-----------|-------------|
| `--input` | JSON file with command array. Use '-' for stdin (NDJSON, one command per line). If omitted, reads from stdin |
| `--session` | Default session ID for all commands. Overridden by per-command sessionId. Auto-captured from session.open/create if not set |
| `--stop-on-error` | Stop execution on first error (default: continue all commands) |
