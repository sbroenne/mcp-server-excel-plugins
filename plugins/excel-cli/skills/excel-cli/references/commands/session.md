### session

Session management. WORKFLOW: open -> use sessionId -> close (--save to persist). Use --show for IRM/auth prompts

#### session create

Create a new Excel file, open it, and create a session. Add --show for visible Excel

| Parameter | Description |
|-----------|-------------|
| `<file>` | Path to the new Excel file to create |
| `--timeout` | Session open/create and operation timeout in whole seconds (default: 120; range: 10-3600) |
| `--show` | Show the Excel window for IRM/auth prompts instead of running hidden |

#### session open

Open an Excel file and create a session. Add --show for visible Excel

| Parameter | Description |
|-----------|-------------|
| `<file>` | Path to the Excel file to open |
| `--timeout` | Session open and operation timeout in whole seconds (default: 120; range: 10-3600) |
| `--show` | Show the Excel window for IRM/auth prompts instead of running hidden |

#### session close

Close a session. Use --save to persist changes

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID to close |
| `--save` | Save changes before closing |

#### session list

List active sessions; transport failures return unresponsive instead of an empty list

#### session test

Test file existence, validity, openability, and IRM/AIP read-only requirements

| Parameter | Description |
|-----------|-------------|
| `<file>` | Full path to test for existence, validity, openability, and IRM/AIP requirements |
| `--timeout` | Excel validation open timeout in whole seconds (default: 120; range: 10-3600) |
