# =============================================================
# clean.ps1 - Remove null bytes from extracted files
# =============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'lib\preflight.ps1')
$ScriptDir = Resolve-ScriptRoot -Invocation $MyInvocation -FallbackBaseDirectory $PSScriptRoot

$bootstrap = Import-CnpjConfig -ScriptRoot $ScriptDir
$config = $bootstrap.Config
$dir = $config.dirTemp

Write-Step 'Cleaning preflight checks'
Invoke-CnpjPreflight -Config $config -EnsureDirectories @($dir) -MinFreeSpaceGB @{ $dir = 8 } -CheckMemory | Out-Null

Write-Host '===================================================' -ForegroundColor Cyan
Write-Host ' CLEAN - Removing null bytes' -ForegroundColor Cyan
Write-Host '===================================================' -ForegroundColor Cyan
Write-Host " Directory: $dir" -ForegroundColor Yellow
Write-Host '-------------------------------------------------' -ForegroundColor Gray

$padroes = @('*.ESTABELE', '*.EMPRECSV', '*.SIMPLES*', '*.CNAECSV', '*.MUNICCSV', '*.NATJUCSV')

$arquivos = @()
foreach ($padrao in $padroes) {
    $encontrados = Get-ChildItem -Path $dir -Filter $padrao -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '^LIMPO_' }
    if ($encontrados) {
        $arquivos += $encontrados
    }
}

if ($arquivos.Count -eq 0) {
    Write-Host 'No raw data files were found.' -ForegroundColor Red
    Write-Host 'Run extract.ps1 first.' -ForegroundColor Red
    exit 1
}

Write-Host "Found $($arquivos.Count) files to check and clean." -ForegroundColor Gray

$cleanupMode = Get-CleanupMode -Config $config
$maxWorkers = if ($config.ContainsKey('cleaningMaxWorkers')) { [Math]::Max(1, [int]$config.cleaningMaxWorkers) } else { 3 }
$rawAllowedPatterns = @('*.ESTABELE', '*.EMPRECSV', '*.SIMPLES*', '*.CNAECSV', '*.MUNICCSV', '*.NATJUCSV')

function Test-CleanOutputValid {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$InputFile,
        [Parameter(Mandatory = $true)][string]$OutputFile,
        [Parameter(Mandatory = $true)][int64]$InputLines,
        [Parameter(Mandatory = $true)][int64]$OutputLines
    )

    if (-not (Test-TextFileReadable -Path $OutputFile)) {
        throw "Cleaned file is missing, empty, or unreadable: $OutputFile"
    }

    if ($InputLines -ne $OutputLines) {
        throw "Line-count mismatch after cleaning $($InputFile.Name): input=$InputLines output=$OutputLines"
    }

    [void](Get-Item -LiteralPath $OutputFile)
}

function Receive-CleanJobResult {
    param(
        [Parameter(Mandatory = $true)][System.Management.Automation.Job]$Job,
        [Parameter(Mandatory = $true)][psobject]$JobInfo
    )

    Wait-Job -Job $Job | Out-Null
    $result = Receive-Job -Job $Job
    Remove-Job -Job $Job -Force
    $record = if ($result -is [System.Array]) { $result[-1] } else { $result }

    if (-not $record.Success) {
        if (Test-Path -LiteralPath $JobInfo.OutputPath) {
            Remove-CnpjArtifactSafe -Config $config -Path $JobInfo.OutputPath -AllowedRoot $dir -AllowedPatterns @('LIMPO_*') -Reason 'failed cleaning output' | Out-Null
        }
        Write-Host "  ERROR in $($JobInfo.InputName): $($record.Error)" -ForegroundColor Red
        return $false
    }

    Test-CleanOutputValid -InputFile $JobInfo.InputFile -OutputFile $JobInfo.OutputPath -InputLines $record.InputLines -OutputLines $record.OutputLines
    if ($cleanupMode -in @('aggressive', 'balanced')) {
        Remove-CnpjArtifactSafe -Config $config -Path $JobInfo.InputPath -AllowedRoot $dir -AllowedPatterns $rawAllowedPatterns -Reason 'cleaned file validated' | Out-Null
    }

    Write-Host ("  OK: {0} ({1} lines)" -f $JobInfo.OutputPath, $record.OutputLines) -ForegroundColor Gray
    return $true
}

$jobs = @()
$erros = 0
$scheduledJobs = 0

foreach ($f in $arquivos) {
    $inputFile = $f.FullName
    $outputFile = Join-Path $dir "LIMPO_$($f.Name)"

    if (Test-Path -LiteralPath $outputFile) {
        if (Test-TextFileReadable -Path $outputFile) {
            if ($cleanupMode -in @('aggressive', 'balanced')) {
                Remove-CnpjArtifactSafe -Config $config -Path $inputFile -AllowedRoot $dir -AllowedPatterns $rawAllowedPatterns -Reason 'existing cleaned file validated' | Out-Null
            }
        }
        Write-Host "[$($arquivos.IndexOf($f) + 1)/$($arquivos.Count)] already cleaned: $($f.Name)" -ForegroundColor Yellow
        continue
    }

    Write-Host "[$($arquivos.IndexOf($f) + 1)/$($arquivos.Count)] Scheduling job: $($f.Name)"

    $job = Start-Job -ScriptBlock {
        param($inputPath, $outputPath)

        $reader = $null
        $writer = $null

        try {
            $reader = [System.IO.File]::OpenText($inputPath)
            $writer = [System.IO.File]::CreateText($outputPath)
            $count = 0

            while ($null -ne ($line = $reader.ReadLine())) {
                if ($line.Contains([char]0)) {
                    $line = $line.Replace([string][char]0, '')
                }

                $writer.WriteLine($line)
                $count++
            }

            [pscustomobject]@{
                Success = $true
                InputPath = $inputPath
                OutputPath = $outputPath
                InputLines = $count
                OutputLines = $count
                BytesWritten = ([System.IO.FileInfo]$outputPath).Length
                Error = $null
            }
        }
        catch {
            [pscustomobject]@{
                Success = $false
                InputPath = $inputPath
                OutputPath = $outputPath
                InputLines = 0
                OutputLines = 0
                BytesWritten = 0
                Error = $_.Exception.Message
            }
        }
        finally {
            if ($reader) { $reader.Dispose() }
            if ($writer) { $writer.Dispose() }
        }
    } -ArgumentList $inputFile, $outputFile

    $scheduledJobs++

    $jobs += [pscustomobject]@{
        Job = $job
        InputFile = $f
        InputName = $f.Name
        InputPath = $inputFile
        OutputPath = $outputFile
    }

    while ($jobs.Count -ge $maxWorkers) {
        $finished = Wait-Job -Job @($jobs.Job) -Any
        $jobInfo = @($jobs | Where-Object { $_.Job.Id -eq $finished.Id })[0]
        if (-not (Receive-CleanJobResult -Job $finished -JobInfo $jobInfo)) { $erros++ }
        $jobs = @($jobs | Where-Object { $_.Job.Id -ne $finished.Id })
    }
}

if ($scheduledJobs -eq 0) {
    Write-Host "`n Cleaning was already 100% complete!" -ForegroundColor Green
    exit 0
}

Write-Host "`nWaiting for $($jobs.Count) remaining cleaning job(s)..." -ForegroundColor Gray
foreach ($jobInfo in @($jobs)) {
    if (-not (Receive-CleanJobResult -Job $jobInfo.Job -JobInfo $jobInfo)) { $erros++ }
}

if ($erros -gt 0) {
    Write-Host "`n[!] WARNING: $erros file(s) failed during cleaning." -ForegroundColor Red
    exit 1
}

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host ' Cleaning completed successfully!' -ForegroundColor Green
