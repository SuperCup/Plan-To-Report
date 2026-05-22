# Publish this project to GitHub (run once after: gh auth login)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $root

$gh = Join-Path $root ".tools\gh\bin\gh.exe"
if (-not (Test-Path $gh)) {
    Write-Host "Missing gh CLI. Download from https://cli.github.com/ or run setup in README."
    exit 1
}

& $gh auth status
if ($LASTEXITCODE -ne 0) {
    Write-Host "Run: gh auth login"
    exit 1
}

$remoteUrl = "https://github.com/SuperCup/Plan-To-Report.git"
if (-not (git remote get-url origin 2>$null)) {
    git remote add origin $remoteUrl
} else {
    git remote set-url origin $remoteUrl
}

git branch -M main
git push -u origin main

Write-Host "Done. Open: https://github.com/SuperCup/Plan-To-Report"
