# Excel CLI Plugin

**Status:** ✅ Phase 1 — Functional with validated skill content. Ready for local testing.

Teaches AI coding agents (Copilot, Cursor, Windsurf) how to use the `excelcli` command-line tool for Excel automation, scripting, and batch workflows.

## Prerequisites

- **Windows** (COM interop required — macOS/Linux unsupported)
- **Microsoft Excel 2016 or later** installed
- **`excelcli.exe` installed and on PATH** (see [Installation](#installation))

## Installation

### Step 1: Download and Install `excelcli.exe`

Choose one method:

**Option A: Direct Download (Recommended)**
1. Download `ExcelMcp-CLI-{version}-windows.zip` from [Releases](https://github.com/sbroenne/mcp-server-excel/releases/latest)
2. Extract `excelcli.exe` to a permanent location (e.g., `C:\Tools\ExcelMcp\`)
3. Add the directory to your PATH

**Option B: NuGet Global Tool (Requires .NET 10 runtime)**
```powershell
dotnet tool install --global Sbroenne.ExcelMcp.CLI
```

**Option C: Package Managers**
```powershell
# Chocolatey (when available)
choco install excelmcp-cli

# Scoop (when available)
scoop install excelmcp-cli
```

**Verify Installation:**
```powershell
excelcli --version
excelcli --help
```

### Step 2: Register the Plugin

```powershell
# Register the plugin marketplace (one-time)
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins

# Install the excel-cli skill
copilot plugin install excel-cli@mcp-server-excel
```

## What's Included

- **230 CLI commands** across 17 categories (Power Query, DAX, PivotTables, Ranges, VBA, etc.)
- **Batch mode** for efficient multi-command workflows
- **Critical rules** for reliable automation (session lifecycle, calculation mode, error handling)
- **Practical examples** for common scenarios (data loading, calculations, formatting)
- **Token-efficient** — uses 64% fewer tokens than MCP Server

## When to Use CLI vs MCP

| Scenario | Tool | Why |
|----------|------|-----|
| **Coding agents** (Copilot, Cursor) | Excel CLI | Single-tool interface, token-efficient |
| **Conversational AI** (Claude Desktop) | Excel MCP | Rich tool discovery, persistent connection |
| **CI/CD pipelines** | Excel CLI | Shell-friendly, exit codes, background daemon |
| **Interactive workflows** | Excel MCP | Natural language, real-time visibility |

## Documentation

- **SKILL.md** (included) — Complete guide with 8 critical rules, batch mode examples, and all 230 commands
- **Source repo** — [Full documentation](https://github.com/sbroenne/mcp-server-excel) with additional examples and troubleshooting

## Support

Report issues and suggest improvements at [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel/issues)

---

**Phase 1 Implementation:** Skill content validated and copied from source repo. Install story clarified with prerequisites and multiple installation methods documented.
