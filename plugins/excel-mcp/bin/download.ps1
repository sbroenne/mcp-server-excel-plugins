# ExcelMcp Binary Download Script
# This script:
# 1. Reads version.txt to get target version
# 2. Downloads mcp-excel.exe from GitHub Release
# 3. Extracts from ZIP
# 4. Reports success/failure

param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$PluginDir = $PSScriptRoot | Split-Path -Parent
$BinaryPath = Join-Path $PluginDir "bin\mcp-excel.exe"
$VersionFile = Join-Path $PluginDir "version.txt"

# Check if binary already exists
if ((Test-Path $BinaryPath) -and -not $Force) {
    Write-Host "✅ MCP Server binary already exists at:" -ForegroundColor Green
    Write-Host "   $BinaryPath"
    Write-Host ""
    Write-Host "To re-download, run with -Force flag:"
    Write-Host "   pwsh -ExecutionPolicy Bypass -File `"$PSScriptRoot\download.ps1`" -Force"
    exit 0
}

# Read target version
if (-not (Test-Path $VersionFile)) {
    Write-Error "❌ version.txt not found at: $VersionFile"
    exit 1
}

$Version = (Get-Content $VersionFile -Raw).Trim()
Write-Host "[Download] Target version: $Version" -ForegroundColor Cyan

# Construct download URL
$RepoOwner = "sbroenne"
$RepoName = "mcp-server-excel"
$AssetName = "ExcelMcp-MCP-Server-$Version-windows.zip"
$DownloadUrl = "https://github.com/$RepoOwner/$RepoName/releases/download/v$Version/$AssetName"

Write-Host "[Download] Downloading from: $DownloadUrl" -ForegroundColor Cyan

# Download ZIP to temp location
$TempZip = Join-Path $env:TEMP "excelmcp-$Version.zip"
$TempExtract = Join-Path $env:TEMP "excelmcp-$Version-extract"

try {
    # Download
    Write-Host "[Download] Downloading binary package..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $TempZip -UseBasicParsing
    
    # Extract
    Write-Host "[Download] Extracting..." -ForegroundColor Yellow
    if (Test-Path $TempExtract) {
        Remove-Item -Recurse -Force $TempExtract
    }
    Expand-Archive -Path $TempZip -DestinationPath $TempExtract
    
    # Find mcp-excel.exe in extracted content
    $ExtractedExe = Get-ChildItem -Path $TempExtract -Recurse -Filter "mcp-excel.exe" | Select-Object -First 1
    
    if (-not $ExtractedExe) {
        Write-Error "❌ mcp-excel.exe not found in downloaded package"
        exit 1
    }
    
    # Copy to bin directory
    Write-Host "[Download] Installing binary..." -ForegroundColor Yellow
    Copy-Item -Path $ExtractedExe.FullName -Destination $BinaryPath -Force
    
    # Verify
    if (Test-Path $BinaryPath) {
        $FileSize = (Get-Item $BinaryPath).Length
        Write-Host ""
        Write-Host "✅ MCP Server binary downloaded successfully!" -ForegroundColor Green
        Write-Host "   Location: $BinaryPath" -ForegroundColor Gray
        Write-Host "   Size: $([math]::Round($FileSize/1MB, 2)) MB" -ForegroundColor Gray
        Write-Host "   Version: $Version" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Next step: Make MCP server globally available" -ForegroundColor Cyan
        Write-Host "   pwsh -ExecutionPolicy Bypass -File `"$PluginDir\bin\install-global.ps1`""
    } else {
        Write-Error "❌ Binary installation failed"
        exit 1
    }
} finally {
    # Cleanup
    if (Test-Path $TempZip) {
        Remove-Item -Force $TempZip
    }
    if (Test-Path $TempExtract) {
        Remove-Item -Recurse -Force $TempExtract
    }
}
