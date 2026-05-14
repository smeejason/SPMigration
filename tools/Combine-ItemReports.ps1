#Requires -Version 5.1
<#
.SYNOPSIS
    Finds all ItemReport_R*.csv files recursively under a starting folder
    and merges them into a single CSV file.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "  SPMT ItemReport Combiner" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

# ── 1. Starting folder ───────────────────────────────────────────────────────

$defaultStart = $PSScriptRoot
if (-not $defaultStart) { $defaultStart = (Get-Location).Path }

$startFolder = Read-Host "Starting folder [$defaultStart]"
if ([string]::IsNullOrWhiteSpace($startFolder)) { $startFolder = $defaultStart }

if (-not (Test-Path -LiteralPath $startFolder -PathType Container)) {
    Write-Host "ERROR: Folder not found: $startFolder" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Searching under: $startFolder" -ForegroundColor Yellow
Write-Host ""

# ── 2. Find matching files ───────────────────────────────────────────────────

$files = @(Get-ChildItem -LiteralPath $startFolder -Recurse -Filter 'ItemReport_R*.csv' |
           Sort-Object FullName)

if ($files.Count -eq 0) {
    Write-Host "No ItemReport_R*.csv files found under that folder." -ForegroundColor Red
    exit 1
}

Write-Host "Found $($files.Count) file(s):" -ForegroundColor Green
$files | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
Write-Host ""

# ── 3. Output path ───────────────────────────────────────────────────────────

$defaultOut  = Join-Path ([Environment]::GetFolderPath('Desktop')) 'ItemReport_Combined.csv'
$outputPath  = Read-Host "Save combined CSV to [$defaultOut]"
if ([string]::IsNullOrWhiteSpace($outputPath)) { $outputPath = $defaultOut }

# Ensure the parent directory exists
$outputDir = Split-Path -Parent $outputPath
if ($outputDir -and -not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

if (Test-Path -LiteralPath $outputPath) {
    $overwrite = Read-Host "File already exists. Overwrite? (Y/N) [Y]"
    if ($overwrite -ne '' -and $overwrite -notmatch '^[Yy]') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        exit 0
    }
}

Write-Host ""

# ── 4. Combine ───────────────────────────────────────────────────────────────

$headerWritten = $false
$totalRows     = 0
$stream        = [System.IO.StreamWriter]::new($outputPath, $false, [System.Text.Encoding]::UTF8)

try {
    foreach ($file in $files) {
        Write-Host "Processing: $($file.Name)" -ForegroundColor Gray -NoNewline

        $reader  = [System.IO.StreamReader]::new($file.FullName, [System.Text.Encoding]::UTF8)
        $fileRows = 0

        try {
            $header = $reader.ReadLine()

            if ($null -eq $header) {
                Write-Host " (empty, skipped)" -ForegroundColor DarkYellow
                continue
            }

            if (-not $headerWritten) {
                $stream.WriteLine($header)
                $headerWritten = $true
            }

            while (-not $reader.EndOfStream) {
                $line = $reader.ReadLine()
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    $stream.WriteLine($line)
                    $fileRows++
                }
            }
        }
        finally {
            $reader.Dispose()
        }

        $totalRows += $fileRows
        Write-Host " — $fileRows row(s)" -ForegroundColor Gray
    }
}
finally {
    $stream.Dispose()
}

# ── 5. Summary ───────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "  Done!" -ForegroundColor Green
Write-Host "  Files combined : $($files.Count)" -ForegroundColor White
Write-Host "  Total data rows: $totalRows" -ForegroundColor White
Write-Host "  Output file    : $outputPath" -ForegroundColor White
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""
