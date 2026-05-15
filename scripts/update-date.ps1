# =============================================================
# update-date.ps1 - Detect and update the latest available dataset
# =============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'lib\preflight.ps1')
$ScriptDir = Resolve-ScriptRoot -Invocation $MyInvocation -FallbackBaseDirectory $PSScriptRoot

$bootstrap = Import-CnpjConfig -ScriptRoot $ScriptDir
$configPath = $bootstrap.ConfigPath
$currentConfig = $bootstrap.Config

Write-Host "Checking Receita Federal dataset index..." -ForegroundColor Cyan

try {
    $html = Invoke-RestMethod -Uri "https://dados-abertos-rf-cnpj.casadosdados.com.br/arquivos/" -UseBasicParsing
}
catch {
    Write-Warn "Could not access the Receita Federal site. Keeping current dataset: $($currentConfig.anoMes)"
    exit 0
}

# Find all links in YYYY-MM-DD/ format
$regex = 'href="(\d{4}-\d{2}-\d{2})/"'
$matches = [regex]::Matches($html, $regex)

if ($matches.Count -eq 0) {
    Write-Warn "No dated folders were found on the site. Keeping current dataset: $($currentConfig.anoMes)"
    exit 0
}

# Sort all dates and keep the newest one
$datas = $matches | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Descending
$latestDate = $datas[0]

Write-Host "Latest dataset found: $latestDate" -ForegroundColor Green

# Update config.ps1
$configContent = Get-Content $configPath -Raw
$regexDate = 'anoMes\s*=\s*"[^"]*"'
$newConfig = $configContent -replace $regexDate, "anoMes       = `"$latestDate`""

if ($configContent -match $regexDate) {
    Set-Content -Path $configPath -Value $newConfig -Encoding UTF8
    Write-Host "config.ps1 updated successfully to: $latestDate" -ForegroundColor Green
} else {
    Write-Host "ERROR: Could not find the 'anoMes' variable in config.ps1" -ForegroundColor Red
    exit 1
}
