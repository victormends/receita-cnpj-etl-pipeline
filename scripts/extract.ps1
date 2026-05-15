# =============================================================
# extract.ps1 - Extract ZIP files using System.IO.Compression
# =============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'lib\preflight.ps1')
$ScriptDir = Resolve-ScriptRoot -Invocation $MyInvocation -FallbackBaseDirectory $PSScriptRoot

$bootstrap = Import-CnpjConfig -ScriptRoot $ScriptDir
$config = $bootstrap.Config

Write-Step 'Extraction preflight checks'
Invoke-CnpjPreflight -Config $config -EnsureDirectories @($config.dirTemp) -MinFreeSpaceGB @{ $config.dirTemp = 12 } -CheckMemory | Out-Null

$zips = @(Get-ChildItem -Path $config.dirDownload -Filter "*.zip")
if ($zips.Count -eq 0) {
    Write-Host "No ZIP files were found. Run download.ps1 first." -ForegroundColor Red
    exit 1
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ExpectedPattern {
    param([string]$ZipBaseName)

    switch -Regex ($ZipBaseName) {
        '^Estabelecimentos' { return 'K3241*.ESTABELE' }
        '^Empresas' { return 'K3241*.EMPRECSV' }
        '^Simples' { return '*.SIMPLES*' }
        '^Municipios' { return '*.MUNICCSV' }
        '^Cnaes' { return '*.CNAECSV' }
        '^Naturezas' { return '*.NATJUCSV' }
        default { return '*' }
    }
}

function Invoke-ZipExtraction {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$Zip,
        [Parameter(Mandatory = $true)][string]$TempDirectory,
        [Parameter(Mandatory = $true)][hashtable]$Config
    )

    $markerFile = Join-Path $TempDirectory "$($Zip.BaseName)_extracted.txt"
    $expectedPattern = Get-ExpectedPattern -ZipBaseName $Zip.BaseName
    $cleanupMode = Get-CleanupMode -Config $Config
    $extractedPaths = New-Object 'System.Collections.Generic.List[string]'

    $zipArchive = $null
    function Close-ZipArchive {
        if ($zipArchive) {
            $zipArchive.Dispose()
            $zipArchive = $null
        }
    }

    function Remove-ZipAfterClose {
        param([string]$Reason)

        Close-ZipArchive
        $lastError = $null
        foreach ($attempt in 1..5) {
            try {
                Remove-CnpjArtifactSafe -Config $Config -Path $Zip.FullName -AllowedRoot $Config.dirDownload -AllowedPatterns @('*.zip') -Reason $Reason | Out-Null
                return
            }
            catch {
                $lastError = $_.Exception.Message
                if ($attempt -lt 5) {
                    Start-Sleep -Seconds 2
                }
            }
        }

        throw "Extraction succeeded, but ZIP cleanup failed after closing the archive: $lastError"
    }

    try {
        if (-not (Test-ZipValid -Path $Zip.FullName)) {
            throw 'ZIP is empty or invalid.'
        }

        $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($Zip.FullName)
        $destinations = @{}

        foreach ($entry in $zipArchive.Entries) {
            if ($entry.FullName.EndsWith('/') -or $entry.Name -eq '') { continue }

            if (-not (Test-FileNameAllowed -FileName $entry.Name -AllowedPatterns @($expectedPattern))) {
                throw "Unexpected file inside $($Zip.Name): $($entry.FullName)"
            }

            $destPath = Assert-PathInsideDirectory -Path (Join-Path $TempDirectory $entry.FullName) -AllowedRoot $TempDirectory
            if ($destinations.ContainsKey($destPath.ToLowerInvariant())) {
                throw "Duplicate extraction destination inside $($Zip.Name): $destPath"
            }
            $destinations[$destPath.ToLowerInvariant()] = $true
            [void]$extractedPaths.Add($destPath)
        }

        if ($extractedPaths.Count -eq 0) {
            throw 'ZIP does not contain an expected data file.'
        }

        $alreadyExtracted = (Test-Path -LiteralPath $markerFile) -and (@($extractedPaths | Where-Object { Test-FileNonEmpty -Path $_ }).Count -eq $extractedPaths.Count)
        if ($alreadyExtracted) {
            if ($cleanupMode -eq 'aggressive') {
                Remove-ZipAfterClose -Reason 'already extracted and validated'
            }

            return [pscustomobject]@{
                Name = $Zip.Name
                Skipped = $true
                Success = $true
                Message = 'already extracted - skipping'
            }
        }

        foreach ($entry in $zipArchive.Entries) {
            if ($entry.FullName.EndsWith('/') -or $entry.Name -eq '') { continue }

            $destPath = Assert-PathInsideDirectory -Path (Join-Path $TempDirectory $entry.FullName) -AllowedRoot $TempDirectory
            $parentPath = Split-Path $destPath -Parent
            if (-not (Test-Path -LiteralPath $parentPath)) {
                New-Object System.IO.DirectoryInfo($parentPath) | ForEach-Object { $_.Create() } | Out-Null
            }

            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destPath, $true)
        }

        foreach ($path in $extractedPaths) {
            if (-not (Test-FileNonEmpty -Path $path)) {
                throw "Extracted file is missing or empty: $path"
            }
        }

        New-Item -ItemType File -Path $markerFile -Force | Out-Null

        if ($cleanupMode -eq 'aggressive') {
            Remove-ZipAfterClose -Reason 'extraction validated'
        }

        return [pscustomobject]@{
            Name = $Zip.Name
            Skipped = $false
            Success = $true
            Message = 'OK'
        }
    }
    catch {
        Remove-Item -LiteralPath $markerFile -Force -ErrorAction SilentlyContinue
        return [pscustomobject]@{
            Name = $Zip.Name
            Skipped = $false
            Success = $false
            Message = $_.Exception.Message
        }
    }
    finally {
        Close-ZipArchive
    }
}

Write-Host "===================================================" -ForegroundColor Cyan
Write-Host " EXTRACT - Extracting ZIP files into temp" -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan

$failed = 0
$completed = 0
foreach ($zip in $zips) {
    $completed++
    Write-Host "[$completed/$($zips.Count)] $($zip.Name)... " -NoNewline
    $result = Invoke-ZipExtraction -Zip $zip -TempDirectory $config.dirTemp -Config $config
    if ($result.Success) {
        if ($result.Skipped) {
            Write-Host $result.Message -ForegroundColor Yellow
        }
        else {
            Write-Host $result.Message -ForegroundColor Green
        }
    }
    else {
        Write-Host "ERROR: $($result.Message)" -ForegroundColor Red
        $failed++
    }
}

if ($failed -gt 0) {
    Write-Host "`n[!] WARNING: $failed ZIP file(s) failed during extraction." -ForegroundColor Red
    exit 1
}

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host ' Extraction completed successfully!' -ForegroundColor Green
