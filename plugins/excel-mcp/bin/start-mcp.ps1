# ExcelMcp MCP Server Wrapper
# This script:
# 1. Detects if mcp-excel.exe is present
# 2. Validates version matches version.txt (optional check)
# 3. Provides clear error messages if binary missing
# 4. Launches the MCP server if all checks pass

$ErrorActionPreference = "Stop"

$PluginDir = $PSScriptRoot | Split-Path -Parent
$BinaryPath = Join-Path $PluginDir "bin\mcp-excel.exe"
$VersionFile = Join-Path $PluginDir "version.txt"

# Check if binary exists
if (-not (Test-Path $BinaryPath)) {
    Write-Error "❌ MCP Server binary not found!"
    Write-Error ""
    Write-Error "The ExcelMcp MCP server binary is missing at:"
    Write-Error "  $BinaryPath"
    Write-Error ""
    Write-Error "To download the binary, run:"
    Write-Error "  pwsh -ExecutionPolicy Bypass -File `"$PluginDir\bin\download.ps1`""
    Write-Error ""
    Write-Error "Or manually download from:"
    Write-Error "  https://github.com/sbroenne/mcp-server-excel/releases"
    exit 1
}

# Optional: Version check (only warn, don't block)
if (Test-Path $VersionFile) {
    $ExpectedVersion = (Get-Content $VersionFile -Raw).Trim()
    # Could add version detection from binary here if needed
    # For now, just log expected version
    Write-Host "[ExcelMcp] Expected version: $ExpectedVersion" -ForegroundColor Gray
}

# Launch the MCP server
Write-Host "[ExcelMcp] Launching MCP server: $BinaryPath" -ForegroundColor Gray

# Pass all args through to the binary
& $BinaryPath @args

# Capture and propagate exit code
$ExitCode = $LASTEXITCODE
exit $ExitCode
