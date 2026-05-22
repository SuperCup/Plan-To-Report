# Zip dist_release for offline distribution (not committed to git by default).
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $root "dist_release\PlanToReport"
if (-not (Test-Path (Join-Path $src "PlanToReport.exe"))) {
    Write-Error "Run build.bat first. Missing dist_release\PlanToReport\PlanToReport.exe"
}
$outDir = Join-Path $root "releases"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$stamp = Get-Date -Format "yyyyMMdd"
$zip = Join-Path $outDir "PlanToReport-win64-$stamp.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path $src -DestinationPath $zip -CompressionLevel Optimal
Write-Host "Release package: $zip"
