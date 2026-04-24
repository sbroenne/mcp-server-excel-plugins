# ExcelMcp Global Install Helper
# This script copies the plugin's MCP server config to user-level MCP config,
# making the MCP server globally available (not just workspace-scoped).
#
# Why needed: Plugin .mcp.json is workspace-scoped by default. For global availability
# across all projects, users need to merge config into ~/.copilot/mcp-config.json.

param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$PluginDir = $PSScriptRoot | Split-Path -Parent
$BinaryPath = Join-Path $PluginDir "bin\mcp-excel.exe"
$UserMcpConfig = Join-Path $env:USERPROFILE ".copilot\mcp-config.json"

Write-Host "ExcelMcp Global Install Helper" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan
Write-Host ""

# Check if binary exists
if (-not (Test-Path $BinaryPath)) {
    Write-Error "❌ MCP Server binary not found!"
    Write-Error ""
    Write-Error "Please download the binary first:"
    Write-Error "   pwsh -ExecutionPolicy Bypass -File `"$PluginDir\bin\download.ps1`""
    exit 1
}

# Ensure ~/.copilot directory exists
$CopilotDir = Join-Path $env:USERPROFILE ".copilot"
if (-not (Test-Path $CopilotDir)) {
    Write-Host "[Install] Creating ~/.copilot directory..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $CopilotDir -Force | Out-Null
}

# Load existing config or create new
$config = @{ mcpServers = @{} }

if (Test-Path $UserMcpConfig) {
    Write-Host "[Install] Loading existing user MCP config..." -ForegroundColor Yellow
    $existingContent = Get-Content $UserMcpConfig -Raw | ConvertFrom-Json
    $config = @{
        mcpServers = if ($existingContent.mcpServers) { 
            $existingContent.mcpServers | ConvertTo-Json -Depth 10 | ConvertFrom-Json 
        } else { 
            @{} 
        }
    }
}

# Check if excel-mcp already exists
if ($config.mcpServers.PSObject.Properties.Name -contains "excel-mcp" -and -not $Force) {
    Write-Host "✅ excel-mcp server is already configured in user MCP config" -ForegroundColor Green
    Write-Host ""
    Write-Host "To reconfigure, run with -Force flag:"
    Write-Host "   pwsh -ExecutionPolicy Bypass -File `"$PSScriptRoot\install-global.ps1`" -Force"
    exit 0
}

# Add excel-mcp server config
Write-Host "[Install] Adding excel-mcp to user MCP config..." -ForegroundColor Yellow

$excelMcpConfig = @{
    command = "powershell"
    args = @(
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        "$PluginDir\bin\start-mcp.ps1"
    )
}

# Convert to PSCustomObject for proper JSON serialization
$config.mcpServers | Add-Member -MemberType NoteProperty -Name "excel-mcp" -Value $excelMcpConfig -Force

# Save to user config
$config | ConvertTo-Json -Depth 10 | Set-Content $UserMcpConfig -Encoding UTF8

Write-Host ""
Write-Host "✅ ExcelMcp MCP server installed globally!" -ForegroundColor Green
Write-Host "   Config: $UserMcpConfig" -ForegroundColor Gray
Write-Host ""
Write-Host "The MCP server is now available in all your projects." -ForegroundColor Cyan
Write-Host ""
Write-Host "Verify installation:" -ForegroundColor Cyan
Write-Host "   copilot mcp list" -ForegroundColor Gray
Write-Host ""
Write-Host "To remove:" -ForegroundColor Cyan
Write-Host "   Edit $UserMcpConfig and remove the 'excel-mcp' entry" -ForegroundColor Gray
