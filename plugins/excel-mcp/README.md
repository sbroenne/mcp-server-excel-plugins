# Excel MCP Server Plugin

**Status:** ✅ Phase 1 — Functional for local testing. Full marketplace deployment blocked by missing GitHub Release.

## Overview

This plugin provides the ExcelMcp MCP Server for conversational AI assistants (Claude, Copilot Chat). The server exposes 25 specialized tools with 230 Excel operations via Model Context Protocol.

**Windows-only** — Requires Microsoft Excel 2016 or later.

**What's Proven (Phase 5 Testing):**
- ✅ Plugin structure spec-compliant
- ✅ All scripts functional (syntax-verified, error-handling tested)
- ✅ Complete skill content bundled
- ✅ Local install/download workflow works with manual binary placement

**What's Blocked:**
- ⚠️ Full E2E workflow requires `v0.0.1` GitHub Release in source repo
- ⚠️ `download.ps1` will fail with 404 until release published

## Installation

### Step 1: Install Plugin

```powershell
# Option A: From local directory (development/testing)
copilot plugin install D:\source\mcp-server-excel-plugins\plugins\excel-mcp

# Option B: From marketplace (when published)
copilot plugin marketplace add sbroenne/mcp-server-excel-plugins
copilot plugin install excel-mcp@mcp-server-excel
```

**⚠️ Important:** The plugin installs successfully, but the MCP server binary is NOT included in the plugin package (it's ~60MB). You must download it separately in Step 2.

### Step 2a: Download MCP Server Binary

After installing the plugin, download the MCP server binary:

```powershell
# Navigate to installed plugin directory
cd $env:USERPROFILE\.copilot\installed-plugins\_direct\excel-mcp\bin

# Run download script
.\download.ps1
```

This downloads `mcp-excel.exe` from the latest GitHub Release for the version specified in `version.txt`.

**⚠️ BLOCKER:** If `download.ps1` fails with a 404 error, the GitHub Release doesn't exist yet. Use the workaround below.

### Step 2b: Workaround (Until Release Published)

If you have a locally-built `mcp-excel.exe` from the source repo development build:

```powershell
# Copy binary directly to plugin directory (skip download.ps1)
Copy-Item "D:\source\mcp-server-excel\artifacts\mcp-excel.exe" `
          "$env:USERPROFILE\.copilot\installed-plugins\_direct\excel-mcp\bin\mcp-excel.exe"
```

Then proceed to Step 2c.

### Step 2c: Make MCP Server Globally Available (Recommended)

By default, plugin MCP servers are **workspace-scoped** — they only work when you're in the plugin directory. To make the MCP server available across ALL your projects:

```powershell
# From the plugin's bin directory
.\install-global.ps1
```

This merges the MCP server config into your user-level `~/.copilot/mcp-config.json`, making it globally available.

### Verify Installation

```powershell
# List all MCP servers
copilot mcp list

# You should see:
# User servers:
#   excel-mcp (local)

# Test wrapper script
cd $env:USERPROFILE\.copilot\installed-plugins\_direct\excel-mcp\bin
.\start-mcp.ps1 --version
# Should output: mcp-excel.exe version X.Y.Z
```

## Why Two Steps?

**Why not bundle the binary in the plugin?**

The MCP server binary (`mcp-excel.exe`) is ~60MB. Including it directly in the plugin would:
- Make plugin installs slow (large download)
- Bloat the plugin repository
- Require repackaging the entire plugin for every release

Instead, the plugin includes a helper script that downloads the binary on-demand from GitHub Releases.

**Current Status:** The source repo hasn't yet published a v0.0.1 GitHub Release, so `download.ps1` will fail until that release is created. This is a **temporary blocker** — use the workaround (manual binary placement) for local testing.

**Why workspace-scoped by default?**

Copilot CLI treats plugin `.mcp.json` files as workspace configurations. This aligns with "project-specific tooling" but requires an extra step for global availability. The `install-global.ps1` script automates this one-time setup.

## What's Included

- **25 specialized tools with 230 MCP operations** — Full Excel automation surface (ranges, tables, Power Query, DAX, PivotTables, charts, VBA, etc.)
- **Excel skill** — Behavioral rules, workflows, and gotchas (19 reference documents)
- **Helper scripts** — Binary download and global install automation

## Phase 1 Status (Completed 2026-04-23)

**Implemented & Validated:**
- ✅ Functional wrapper script (`start-mcp.ps1`) — error handling for missing binary verified
- ✅ Functional download script (`download.ps1`) — logic validated, URL construction correct
- ✅ Functional global install helper (`install-global.ps1`) — config merge logic validated
- ✅ Real skill content from source repo (25 tools with 230 operations documented)
- ✅ Plugin manifest spec-compliant
- ✅ MCP config uses `{pluginDir}` placeholder (proven in earlier spike)
- ✅ Local testing workflow: install plugin → manually place binary → test

**Tested Locally:**
- ✅ All three scripts syntax-validated (PowerShell parser)
- ✅ Error handling verified (missing binary scenario)
- ✅ Version detection from `version.txt` confirmed

**Blocked (Requires Source Repo Action):**
- ⚠️ Full E2E install → download → global-install → mcp list workflow NOT YET TESTED
  - **Reason:** No v0.0.1 GitHub Release published in source repo
  - **Impact:** `download.ps1` will fail with 404 error
  - **Workaround:** Manually place binary, skip download step

**Next Steps for Full Validation:**
1. Source repo publishes v0.0.1 GitHub Release with `ExcelMcp-MCP-Server-0.0.1-windows.zip`
2. Test full workflow: install → download → global-install → verify MCP server
3. Then promote plugin to published marketplace

## Local Testing

To test the plugin locally before marketplace publication:

```powershell
# 1. Install from local directory
copilot plugin install D:\source\mcp-server-excel-plugins\plugins\excel-mcp

# 2. Manually place mcp-excel.exe in plugin's bin/ directory
#    (skip download.ps1 if no release exists yet)

# 3. Test wrapper script
powershell -ExecutionPolicy Bypass -File $env:USERPROFILE\.copilot\installed-plugins\_direct\excel-mcp\bin\start-mcp.ps1 --version

# 4. Install globally
cd $env:USERPROFILE\.copilot\installed-plugins\_direct\excel-mcp\bin
.\install-global.ps1

# 5. Verify
copilot mcp list
```

## Documentation

See [source repo documentation](https://github.com/sbroenne/mcp-server-excel) for full details on:
- MCP Server tool reference
- Excel COM interop patterns
- Power Query (M code) syntax
- DAX measure patterns
- Integration test coverage

## Support

Report issues at [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel/issues)

---

**Phase 1 implementation by:** Kelso (Copilot CLI Plugin Engineer)  
**Completed in Phase 1:**
- Functional scripts (wrapper, download, global install)
- Real skill content copied from source repo
- Accurate install docs reflecting workspace-scoped reality
- Spec-compliant plugin.json and .mcp.json
