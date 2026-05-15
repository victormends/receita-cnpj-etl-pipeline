# =============================================================
# download.ps1 - Download all source ZIP files
# =============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'lib\preflight.ps1')
$ScriptDir = Resolve-ScriptRoot -Invocation $MyInvocation -FallbackBaseDirectory $PSScriptRoot

$bootstrap = Import-CnpjConfig -ScriptRoot $ScriptDir
$config = $bootstrap.Config

Write-Step 'Download preflight checks'
Invoke-CnpjPreflight -Config $config -EnsureDirectories @($config.dirDownload) -MinFreeSpaceGB @{ $config.dirDownload = 8 } | Out-Null

$arquivos = @(
    'Cnaes.zip', 'Municipios.zip', 'Naturezas.zip', 'Simples.zip',
    'Estabelecimentos0.zip', 'Estabelecimentos1.zip', 'Estabelecimentos2.zip',
    'Estabelecimentos3.zip', 'Estabelecimentos4.zip', 'Estabelecimentos5.zip',
    'Estabelecimentos6.zip', 'Estabelecimentos7.zip', 'Estabelecimentos8.zip',
    'Estabelecimentos9.zip',
    'Empresas0.zip', 'Empresas1.zip', 'Empresas2.zip', 'Empresas3.zip',
    'Empresas4.zip', 'Empresas5.zip', 'Empresas6.zip', 'Empresas7.zip',
    'Empresas8.zip', 'Empresas9.zip'
)

$baseUrl = "$($config.baseUrl)/$($config.anoMes)"

Write-Host '===================================================' -ForegroundColor Cyan
Write-Host " DOWNLOAD - CNPJ files for $($config.anoMes)" -ForegroundColor Cyan
Write-Host '===================================================' -ForegroundColor Cyan
Write-Host " URL Base: $baseUrl" -ForegroundColor Yellow
Write-Host " Destino: $($config.dirDownload)" -ForegroundColor Yellow
Write-Host '-------------------------------------------------' -ForegroundColor Gray

$failed = @()

foreach ($arquivo in $arquivos) {
    $url = "$baseUrl/$arquivo"
    $destino = Join-Path $config.dirDownload $arquivo
    $destinoTmp = "$destino.partial"

    Write-Host "[$($arquivos.IndexOf($arquivo) + 1)/$($arquivos.Count)] $arquivo..." -NoNewline

    if (Test-Path -LiteralPath $destino) {
        $existingFile = Get-Item $destino
        if (Test-ZipValid -Path $destino) {
            $size = [math]::Round($existingFile.Length / 1MB, 1)
            Write-Host " valid ZIP already exists ($size MB) - skipping" -ForegroundColor Yellow
            continue
        }

        Write-Host " invalid existing ZIP - deleting and retrying..." -ForegroundColor Yellow
        Remove-CnpjArtifactSafe -Config $config -Path $destino -AllowedRoot $config.dirDownload -AllowedPatterns @('*.zip') -Reason 'invalid existing download' | Out-Null
    }

    if (Test-Path -LiteralPath $destinoTmp) {
        Remove-CnpjArtifactSafe -Config $config -Path $destinoTmp -AllowedRoot $config.dirDownload -AllowedPatterns @('*.partial') -Reason 'stale partial download' | Out-Null
    }

    try {
        $start = Get-Date
        Invoke-WebRequest -Uri $url -OutFile $destinoTmp -UseBasicParsing

        if (-not (Test-Path -LiteralPath $destinoTmp)) {
            throw 'Download finished without creating a temporary file.'
        }

        if (-not (Test-ZipValid -Path $destinoTmp)) {
            throw 'Downloaded file is not a valid ZIP archive.'
        }

        Move-Item $destinoTmp $destino -Force

        $duration = ((Get-Date) - $start).TotalSeconds
        $size = [math]::Round(((Get-Item $destino).Length / 1MB), 1)
        $speed = if ($duration -gt 0) { [math]::Round($size / $duration, 1) } else { 0 }
        Write-Host " OK ($size MB, $speed MB/s)" -ForegroundColor Green
    }
    catch {
        Write-Host " ERROR: $($_.Exception.Message)" -ForegroundColor Red
        if (Test-Path -LiteralPath $destinoTmp) {
            Remove-CnpjArtifactSafe -Config $config -Path $destinoTmp -AllowedRoot $config.dirDownload -AllowedPatterns @('*.partial') -Reason 'failed download partial' | Out-Null
        }
        $failed += $arquivo
    }
}

if ($failed.Count -gt 0) {
    Write-Host "`n[!] WARNING: Some files failed to download:" -ForegroundColor Red
    $failed | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    Write-Host "`nRun the script again. It will skip valid ZIPs and retry failed files." -ForegroundColor Yellow
    exit 1
}

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host ' Download completed!' -ForegroundColor Green
