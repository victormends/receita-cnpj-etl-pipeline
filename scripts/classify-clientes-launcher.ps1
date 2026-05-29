[CmdletBinding()]
param(
    [string]$InputPath,
    [string]$SimplesPath,
    [string]$OutputPath,
    [string]$EnrichmentPath,
    [string]$ClassifiedInputPath,
    [string]$XmlRoot,
    [string]$Delimiter = ';',
    [ValidateSet('Database', 'File')]
    [string]$Mode = 'Database'
)

# =============================================================
# classify-clientes-launcher.ps1 - Interactive launcher for the client classifier
# =============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Wait-BeforeExit {
    Write-Host ''
    Write-Host 'Press Enter to close.' -ForegroundColor Yellow
    [void][Console]::ReadLine()
}

function Read-ConsoleValue {
    param([Parameter(Mandatory = $true)][string]$Prompt)

    Write-Host -NoNewline "$Prompt`: "
    $value = [Console]::ReadLine()
    if ($null -eq $value) { return '' }
    return $value
}

function Select-OpenFile {
    param([string]$Title)

    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = 'Excel or CSV files (*.xlsx;*.csv)|*.xlsx;*.csv|Excel files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $dialog.CheckFileExists = $true
    $dialog.Multiselect = $false

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return ''
    }

    return $dialog.FileName
}

function Select-SaveFile {
    param([string]$Title, [string]$DefaultFileName = 'clientes_classificados.csv')

    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = $Title
    $dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $dialog.FileName = $DefaultFileName
    $dialog.OverwritePrompt = $true

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return ''
    }

    return $dialog.FileName
}

function Read-Delimiter {
    param([string]$Default = ';')

    Add-Type -AssemblyName Microsoft.VisualBasic
    $value = [Microsoft.VisualBasic.Interaction]::InputBox('CSV delimiter:', 'CNPJ Client Classifier', $Default)
    if ([string]::IsNullOrWhiteSpace($value)) { return $Default }
    return $value
}

function Get-LauncherDirectory {
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        return $PSScriptRoot
    }

    if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        return (Split-Path -Parent $PSCommandPath)
    }

    if ($MyInvocation.MyCommand.PSObject.Properties.Name -contains 'Path' -and -not [string]::IsNullOrWhiteSpace($MyInvocation.MyCommand.Path)) {
        return (Split-Path -Parent $MyInvocation.MyCommand.Path)
    }

    $entryPath = [Environment]::GetCommandLineArgs()[0]
    if (-not [string]::IsNullOrWhiteSpace($entryPath) -and (Test-Path -LiteralPath $entryPath -PathType Leaf)) {
        return (Split-Path -Parent (Resolve-Path -LiteralPath $entryPath).Path)
    }

    $processPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    if (-not [string]::IsNullOrWhiteSpace($processPath) -and (Test-Path -LiteralPath $processPath -PathType Leaf)) {
        return (Split-Path -Parent $processPath)
    }

    return (Get-Location).Path
}

try {
    $launchedWithArguments = -not [string]::IsNullOrWhiteSpace($InputPath) -or
        -not [string]::IsNullOrWhiteSpace($SimplesPath) -or
        -not [string]::IsNullOrWhiteSpace($OutputPath)

    $scriptDir = Get-LauncherDirectory
    $classifier = Join-Path $scriptDir 'classify-clientes.ps1'
    $stagingBuilder = Join-Path $scriptDir 'build-client-staging.ps1'

    if (-not (Test-Path -LiteralPath $classifier -PathType Leaf)) {
        throw "Classifier script was not found beside the launcher: $classifier"
    }

    Write-Host '===================================================' -ForegroundColor Cyan
    Write-Host ' CLIENT STAGING BUILDER' -ForegroundColor Cyan
    Write-Host '===================================================' -ForegroundColor Cyan
    Write-Host 'Provide the active clients XLSX/CSV. The bundled cleaned enrichment CSV is used automatically when present.' -ForegroundColor Yellow
    Write-Host 'If Downloads\clientes_classificados.csv exists, it is used automatically for MEI/Simples/Normal.' -ForegroundColor Yellow
    Write-Host 'Government CNAE/categories are attached from PostgreSQL when available.' -ForegroundColor Yellow
    Write-Host '-------------------------------------------------' -ForegroundColor Gray

    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        $InputPath = Select-OpenFile 'Select the client XLSX or CSV file'
    }

    if ([string]::IsNullOrWhiteSpace($ClassifiedInputPath)) {
        $defaultClassified = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads\clientes_classificados.csv'
        if (Test-Path -LiteralPath $defaultClassified -PathType Leaf) {
            $ClassifiedInputPath = $defaultClassified
        }
    }

    if ([string]::IsNullOrWhiteSpace($SimplesPath) -and [string]::IsNullOrWhiteSpace($ClassifiedInputPath)) {
        $SimplesPath = Select-OpenFile 'Select the Receita Simples CSV file, or cancel to skip regime classification'
    }

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = Select-SaveFile 'Choose the normalized PostgreSQL output CSV file' 'clientes_postgres_normalizado.csv'
    }

    if ([string]::IsNullOrWhiteSpace($Delimiter)) {
        $Delimiter = Read-Delimiter ';'
    }

    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        throw 'Client CSV path is required.'
    }

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        throw 'Output CSV path is required.'
    }

    if (Test-Path -LiteralPath $stagingBuilder -PathType Leaf) {
        $builderArgs = @{
            InputPath = $InputPath
            OutputPath = $OutputPath
            Delimiter = $Delimiter
            Mode = $Mode
        }
        if (-not [string]::IsNullOrWhiteSpace($SimplesPath)) { $builderArgs.SimplesPath = $SimplesPath }
        if (-not [string]::IsNullOrWhiteSpace($EnrichmentPath)) { $builderArgs.EnrichmentPath = $EnrichmentPath }
        if (-not [string]::IsNullOrWhiteSpace($ClassifiedInputPath)) { $builderArgs.ClassifiedInputPath = $ClassifiedInputPath }
        if (-not [string]::IsNullOrWhiteSpace($XmlRoot)) { $builderArgs.XmlRoot = $XmlRoot }
        $builderArgs.IncludeGovernmentData = $true
        & $stagingBuilder @builderArgs
    }
    else {
        $classifierArgs = @{
            InputPath = $InputPath
            OutputPath = $OutputPath
            Delimiter = $Delimiter
            Mode = $Mode
        }
        if (-not [string]::IsNullOrWhiteSpace($SimplesPath)) { $classifierArgs.SimplesPath = $SimplesPath }
        if (-not [string]::IsNullOrWhiteSpace($XmlRoot)) { $classifierArgs.XmlRoot = $XmlRoot }
        & $classifier @classifierArgs
    }
}
catch {
    Write-Host ''
    Write-Host 'ERROR' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
finally {
    if (-not $launchedWithArguments) {
        Wait-BeforeExit
    }
}
