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

$repoName = "Plan-To-Report"
$exists = & $gh repo view $repoName 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "Remote repo already exists. Pushing..."
    if (-not (git remote get-url origin 2>$null)) {
        $user = (& $gh api user -q .login)
        git remote add origin "https://github.com/$user/$repoName.git"
    }
} else {
    & $gh repo create $repoName --public --source=. --remote=origin --description "Excel plan sheet to activity summary and UPC report tool"
}

git branch -M main
git push -u origin main

Write-Host "Done. Open: https://github.com/$( & $gh api user -q .login )/$repoName"
