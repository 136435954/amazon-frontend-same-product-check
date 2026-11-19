param(
    [string]$Dest = ""
)

$ErrorActionPreference = "Stop"
$skillName = "amazon-frontend-same-product-check"
$src = Join-Path $PSScriptRoot "skill"

if (-not (Test-Path $src)) {
    Write-Host "[x] skill folder not found: $src" -ForegroundColor Red
    exit 1
}

if (-not $Dest) {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA "Doubao\User Data\Profile 1\.doubao\agent_mode\workspace\.user_skills"),
        (Join-Path $env:LOCALAPPDATA "Doubao\User Data\Profile 1\.doubao\agent_mode\workspace\.skills")
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { $Dest = $c; break }
    }
    if (-not $Dest) {
        $hit = Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA "Doubao") -Recurse -Directory -Filter "user_skills" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($hit) { $Dest = $hit.FullName }
    }
}

if (-not $Dest) {
    Write-Host "[x] .user_skills not found. Pass the folder explicitly:" -ForegroundColor Red
    Write-Host '  .\install.ps1 -Dest "C:\Users\<you>\AppData\Local\Doubao\User Data\Profile 1\.doubao\agent_mode\workspace\.user_skills"'
    exit 1
}

$target = Join-Path $Dest $skillName
if (Test-Path $target) {
    Write-Host "[!] existing skill will be replaced: $target" -ForegroundColor Yellow
    Remove-Item -LiteralPath $target -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $target | Out-Null
Copy-Item -Path (Join-Path $src "*") -Destination $target -Recurse -Force

$count = (Get-ChildItem $target -Recurse -File | Measure-Object).Count
Write-Host "[ok] installed $count files to: $target" -ForegroundColor Green
Write-Host "     Restart the chat so the skill gets loaded."
Write-Host ""
Write-Host "Self-check (no network, no tokens):"
Write-Host ('  python "' + $target + '\scripts\amazon_frontend_check.py" --workbook <xlsx> --site US --target <image> --print-config')
