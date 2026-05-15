Set-StrictMode -Version Latest

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = New-Object System.Text.UTF8Encoding $false
}
catch {
}

function Write-Step {
    param([string]$Message)
    Write-Host "`n[>] $Message" -ForegroundColor Cyan
}

function Write-OK {
    param([string]$Message)
    Write-Host "    [OK] $Message" -ForegroundColor Green
}

function Write-Info {
    param([string]$Message)
    Write-Host "    [INFO] $Message" -ForegroundColor Gray
}

function Write-Warn {
    param([string]$Message)
    Write-Host "    [WARN] $Message" -ForegroundColor Yellow
}

function Write-Fail {
    param([string]$Message)
    Write-Host "    [FAIL] $Message" -ForegroundColor Red
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-CurrentLauncherPath {
    if ($MyInvocation.MyCommand -and $MyInvocation.MyCommand.PSObject.Properties['Path'] -and -not [string]::IsNullOrWhiteSpace([string]$MyInvocation.MyCommand.Path)) {
        return [string]$MyInvocation.MyCommand.Path
    }

    $commandLineArgs = [Environment]::GetCommandLineArgs()
    if ($commandLineArgs.Length -gt 0 -and -not [string]::IsNullOrWhiteSpace($commandLineArgs[0])) {
        return [string]$commandLineArgs[0]
    }

    return $null
}

function Restart-CurrentProcessElevated {
    param([string]$Reason)

    $launcherPath = Get-CurrentLauncherPath
    if ([string]::IsNullOrWhiteSpace($launcherPath)) {
        throw "Administrator privileges are required, but the current launcher path could not be determined. Reason: $Reason"
    }

    Write-Warn "Administrator privileges are required: $Reason"
    Write-Info 'Requesting elevation...'

    $startInfo = @{ FilePath = $launcherPath; Verb = 'RunAs' }
    $extension = [System.IO.Path]::GetExtension($launcherPath)
    if ($extension -ieq '.ps1') {
        $startInfo = @{
            FilePath = 'powershell.exe'
            ArgumentList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $launcherPath)
            Verb = 'RunAs'
            WorkingDirectory = (Split-Path -Parent $launcherPath)
        }
    }
    else {
        $startInfo['WorkingDirectory'] = (Split-Path -Parent $launcherPath)
    }

    try {
        Start-Process @startInfo | Out-Null
    }
    catch {
        throw "Failed to relaunch elevated. Reason: $Reason. $($_.Exception.Message)"
    }

    exit 0
}

function Resolve-ScriptRoot {
    param(
        $Invocation = $null,
        [string]$FallbackBaseDirectory = $null
    )

    $trimmedFallback = if ([string]::IsNullOrWhiteSpace($FallbackBaseDirectory)) {
        $null
    }
    else {
        $FallbackBaseDirectory.TrimEnd('\\')
    }

    if ($Invocation) {
        $command = $Invocation.MyCommand
        if ($command -and $command.PSObject.Properties['Path'] -and -not [string]::IsNullOrWhiteSpace([string]$command.Path)) {
            return (Split-Path -Parent $command.Path)
        }
    }

    if ($trimmedFallback) {
        return $trimmedFallback
    }

    if ($PSCommandPath) {
        return (Split-Path -Parent $PSCommandPath)
    }

    if ($PSScriptRoot) {
        return $PSScriptRoot
    }

    if ([System.AppDomain]::CurrentDomain.BaseDirectory) {
        return [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\\')
    }

    throw 'Unable to resolve the script root.'
}

function Get-ProjectRoot {
    param([Parameter(Mandatory = $true)][string]$ScriptRoot)
    return (Split-Path -Parent $ScriptRoot)
}

function Import-CnpjConfig {
    param([Parameter(Mandatory = $true)][string]$ScriptRoot)

    $projectRoot = Get-ProjectRoot -ScriptRoot $ScriptRoot
    $configPath = Join-Path $projectRoot 'config.ps1'
    if (-not (Test-Path $configPath)) {
        throw "config.ps1 not found. Copy config.example.ps1 to config.ps1 and customize it before running. Expected path: $configPath"
    }

    $config = & $configPath
    if (-not ($config -is [hashtable])) {
        throw 'config.ps1 did not return a valid hashtable.'
    }

    $config = Resolve-DatabaseConfig -Config $config
    $config = Resolve-PortableWorkingDirectories -Config $config

    return @{
        ProjectRoot = $projectRoot
        ConfigPath = $configPath
        Config = $config
    }
}

function Get-EnvConfigValue {
    param(
        [Parameter(Mandatory = $true)][string[]]$Names,
        [string]$Default = $null
    )

    foreach ($name in $Names) {
        $value = [Environment]::GetEnvironmentVariable($name)
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }

    return $Default
}

function Resolve-DatabaseConfig {
    param([Parameter(Mandatory = $true)][hashtable]$Config)

    $resolved = @{}
    foreach ($entry in $Config.GetEnumerator()) {
        $resolved[$entry.Key] = $entry.Value
    }

    $defaultDbHost = if ($resolved.ContainsKey('dbHost')) { $resolved['dbHost'] } else { $null }
    $defaultDbPort = if ($resolved.ContainsKey('dbPort')) { $resolved['dbPort'] } else { $null }
    $defaultDbUser = if ($resolved.ContainsKey('dbUser')) { $resolved['dbUser'] } else { $null }
    $defaultDbPassword = if ($resolved.ContainsKey('dbPassword')) { $resolved['dbPassword'] } else { $null }

    $resolved['dbHost'] = Get-EnvConfigValue -Names @('CNPJ_ETL_DB_HOST', 'PGHOST') -Default $defaultDbHost
    $resolved['dbPort'] = Get-EnvConfigValue -Names @('CNPJ_ETL_DB_PORT', 'PGPORT') -Default $defaultDbPort
    $resolved['dbUser'] = Get-EnvConfigValue -Names @('CNPJ_ETL_DB_USER', 'PGUSER') -Default $defaultDbUser
    $resolved['dbPassword'] = Get-EnvConfigValue -Names @('CNPJ_ETL_DB_PASSWORD', 'PGPASSWORD') -Default $defaultDbPassword

    return $resolved
}

function Resolve-PortableWorkingDirectories {
    param([Parameter(Mandatory = $true)][hashtable]$Config)

    $resolved = @{}
    foreach ($entry in $Config.GetEnumerator()) {
        $resolved[$entry.Key] = $entry.Value
    }

    $postgresRoots = @('C:\Postgres17', 'D:\Postgres17', 'C:\Postgres16', 'D:\Postgres16')
    $availablePostgresRoot = $postgresRoots | Where-Object { Test-Path $_ } | Select-Object -First 1

    foreach ($key in @('dirTemp', 'dirOut')) {
        if (-not $resolved.ContainsKey($key)) {
            continue
        }

        $pathValue = [string]$resolved[$key]
        if ([string]::IsNullOrWhiteSpace($pathValue)) {
            continue
        }

        $qualifier = Split-Path -Path $pathValue -Qualifier
        if ($qualifier -and -not (Test-Path $qualifier) -and $availablePostgresRoot) {
            $leaf = Split-Path -Path $pathValue -Leaf
            $resolved[$key] = Join-Path $availablePostgresRoot (Join-Path 'data' $leaf)
        }
    }

    return $resolved
}

function Test-ConfigKey {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [Parameter(Mandatory = $true)][string]$Key
    )

    return $Config.ContainsKey($Key) -and -not [string]::IsNullOrWhiteSpace([string]$Config[$Key])
}

function Assert-ConfigValid {
    param([Parameter(Mandatory = $true)][hashtable]$Config)

    $requiredKeys = @(
        'anoMes', 'baseUrl', 'dirDownload', 'dirTemp', 'dirOut',
        'ufs', 'cnaeWhitelist', 'cnaeRanges',
        'dbHost', 'dbName'
    )

    foreach ($key in $requiredKeys) {
        if (-not (Test-ConfigKey -Config $Config -Key $key)) {
            throw "Required config key is missing or empty: $key"
        }
    }

    if ($Config.anoMes -notmatch '^\d{4}-\d{2}(-\d{2})?$') {
        throw "Invalid anoMes value: '$($Config.anoMes)'. Use YYYY-MM or YYYY-MM-DD."
    }

    if (-not $Config.ContainsKey('dataCorte')) {
        throw 'Required config key is missing: dataCorte'
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$Config.dataCorte) -and $Config.dataCorte -notmatch '^\d{8}$') {
        throw "Invalid dataCorte value: '$($Config.dataCorte)'. Use YYYYMMDD or leave it empty for last 6 months."
    }

    if ((Test-ConfigKey -Config $Config -Key 'dbPort') -and [string]$Config['dbPort'] -notmatch '^\d+$') {
        throw "Invalid dbPort value: '$($Config.dbPort)'."
    }

    if ($Config.ufs.Count -eq 0) {
        throw 'Config ufs cannot be empty.'
    }

    if ($Config.ContainsKey('municipiosWhitelist')) {
        foreach ($municipio in @($Config.municipiosWhitelist)) {
            if ([string]$municipio -notmatch '^\d{4}$') {
                throw "Invalid municipiosWhitelist value: '$municipio'. Use Receita municipality codes with 4 digits."
            }
        }
    }

    if ($Config.ContainsKey('cleanupMode') -and [string]$Config.cleanupMode -notin @('aggressive', 'balanced', 'debug')) {
        throw "Invalid cleanupMode value: '$($Config.cleanupMode)'. Use aggressive, balanced, or debug."
    }
}

function Get-CleanupMode {
    param([Parameter(Mandatory = $true)][hashtable]$Config)

    if ($Config.ContainsKey('cleanupMode') -and -not [string]::IsNullOrWhiteSpace([string]$Config.cleanupMode)) {
        return [string]$Config.cleanupMode
    }

    return 'aggressive'
}

function Test-CleanupDryRun {
    param([Parameter(Mandatory = $true)][hashtable]$Config)

    return ($Config.ContainsKey('cleanupDryRun') -and [bool]$Config.cleanupDryRun)
}

function Get-FullPathSafe {
    param([Parameter(Mandatory = $true)][string]$Path)
    return [System.IO.Path]::GetFullPath($Path)
}

function Assert-PathInsideDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$AllowedRoot
    )

    $fullPath = Get-FullPathSafe -Path $Path
    $fullRoot = (Get-FullPathSafe -Path $AllowedRoot).TrimEnd('\') + '\'
    if (-not $fullPath.StartsWith($fullRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing unsafe path outside allowed root. Path: $fullPath Root: $fullRoot"
    }

    return $fullPath
}

function Test-FileNonEmpty {
    param([Parameter(Mandatory = $true)][string]$Path)

    return (Test-Path -LiteralPath $Path -PathType Leaf) -and ((Get-Item -LiteralPath $Path).Length -gt 0)
}

function Test-ZipValid {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-FileNonEmpty -Path $Path)) {
        return $false
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        return (@($archive.Entries | Where-Object { -not $_.FullName.EndsWith('/') -and $_.Name -ne '' }).Count -gt 0)
    }
    catch {
        return $false
    }
    finally {
        if ($archive) { $archive.Dispose() }
    }
}

function Test-TextFileReadable {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-FileNonEmpty -Path $Path)) {
        return $false
    }

    $reader = $null
    try {
        $reader = [System.IO.File]::OpenText($Path)
        [void]$reader.ReadLine()
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($reader) { $reader.Dispose() }
    }
}

function Test-FileNameAllowed {
    param(
        [Parameter(Mandatory = $true)][string]$FileName,
        [Parameter(Mandatory = $true)][string[]]$AllowedPatterns
    )

    foreach ($pattern in $AllowedPatterns) {
        if ($FileName -like $pattern) {
            return $true
        }
    }

    return $false
}

function Remove-CnpjArtifactSafe {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$AllowedRoot,
        [Parameter(Mandatory = $true)][string[]]$AllowedPatterns,
        [Parameter(Mandatory = $true)][string]$Reason
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{ Deleted = $false; DryRun = $false; Bytes = 0; Path = $Path; Reason = 'not found' }
    }

    $fullPath = Assert-PathInsideDirectory -Path $Path -AllowedRoot $AllowedRoot
    $item = Get-Item -LiteralPath $fullPath
    if (-not (Test-FileNameAllowed -FileName $item.Name -AllowedPatterns $AllowedPatterns)) {
        throw "Refusing to delete file with disallowed name: $($item.FullName)"
    }

    $bytes = $item.Length
    if (Test-CleanupDryRun -Config $Config) {
        Write-Info "DRY RUN cleanup would delete: $fullPath ($([math]::Round($bytes / 1MB, 2)) MB). Reason: $Reason"
        return [pscustomobject]@{ Deleted = $false; DryRun = $true; Bytes = $bytes; Path = $fullPath; Reason = $Reason }
    }

    Remove-Item -LiteralPath $fullPath -Force -ErrorAction Stop
    Write-Info "Cleanup deleted: $fullPath ($([math]::Round($bytes / 1MB, 2)) MB). Reason: $Reason"
    return [pscustomobject]@{ Deleted = $true; DryRun = $false; Bytes = $bytes; Path = $fullPath; Reason = $Reason }
}

function ConvertTo-PsqlPathLiteral {
    param([Parameter(Mandatory = $true)][string]$Path)

    return ($Path -replace '\\', '/') -replace "'", "''"
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $property = $Object.PSObject.Properties[$Name]
    if (-not $property) {
        return $null
    }

    return $property.Value
}

function Get-PostgresRegistryInstallations {
    $roots = @(
        'HKLM:\SOFTWARE\PostgreSQL\Installations',
        'HKLM:\SOFTWARE\WOW6432Node\PostgreSQL\Installations'
    )

    $installs = @()
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) {
            continue
        }

        foreach ($item in (Get-ChildItem -Path $root -ErrorAction SilentlyContinue)) {
            $props = Get-ItemProperty -Path $item.PSPath -ErrorAction SilentlyContinue
            if (-not $props) {
                continue
            }

            $baseDirectory = Get-ObjectPropertyValue -Object $props -Name 'Base Directory'
            if ([string]::IsNullOrWhiteSpace($baseDirectory)) {
                continue
            }

            $serviceId = Get-ObjectPropertyValue -Object $props -Name 'ServiceID'
            $version = Get-ObjectPropertyValue -Object $props -Name 'Version'
            $installs += @{
                BaseDirectory = $baseDirectory.TrimEnd('\\')
                DataDirectory = Get-ObjectPropertyValue -Object $props -Name 'Data Directory'
                Port = Get-ObjectPropertyValue -Object $props -Name 'Port'
                Service = if ($serviceId) { [string]$serviceId } else { $null }
                Version = if ($version) { [string]$version } else { $null }
            }
        }
    }

    return $installs
}

function Get-PostgresPrefixCandidates {
    $prefixes = New-Object System.Collections.Generic.List[string]
    $seen = @{}

    function Add-PrefixCandidate {
        param([string]$Path)

        if ([string]::IsNullOrWhiteSpace($Path)) {
            return
        }

        $normalized = $Path.TrimEnd('\\')
        if ($seen.ContainsKey($normalized)) {
            return
        }

        $seen[$normalized] = $true
        $prefixes.Add($normalized)
    }

    foreach ($path in @(
        'D:\Postgres17',
        'C:\Postgres17',
        'D:\Postgres16',
        'C:\Postgres16',
        'D:\Postgres15',
        'C:\Postgres15',
        'C:\Program Files\PostgreSQL',
        'C:\Program Files (x86)\PostgreSQL',
        'C:\PostgreSQL',
        'D:\PostgreSQL'
    )) {
        Add-PrefixCandidate -Path $path
    }

    foreach ($install in (Get-PostgresRegistryInstallations)) {
        Add-PrefixCandidate -Path $install.BaseDirectory
    }

    return $prefixes.ToArray()
}

function Resolve-ExecutablePath {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [string[]]$PreferredPrefixes = @()
    )

    foreach ($prefix in $PreferredPrefixes) {
        if ([string]::IsNullOrWhiteSpace($prefix)) {
            continue
        }

        $candidate = Join-Path $prefix (Join-Path 'bin' $Name)
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    $command = Get-Command $Name -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) {
        return $command.Source
    }

    foreach ($basePath in (Get-PostgresPrefixCandidates)) {
        if (-not (Test-Path $basePath)) { continue }

        $directCandidate = Join-Path $basePath (Join-Path 'bin' $Name)
        if (Test-Path $directCandidate) {
            return $directCandidate
        }

        $versionFolders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending
        foreach ($folder in $versionFolders) {
            $candidate = Join-Path $folder.FullName (Join-Path 'bin' $Name)
            if (Test-Path $candidate) {
                return $candidate
            }
        }
    }

    return $null
}

function Resolve-PostgresTooling {
    param([Parameter(Mandatory = $true)][hashtable]$Config)

    $registryInstalls = Get-PostgresRegistryInstallations
    $preferredPrefixes = New-Object System.Collections.Generic.List[string]
    foreach ($install in $registryInstalls) {
        if ($install.BaseDirectory -and (Test-Path $install.BaseDirectory)) {
            $preferredPrefixes.Add($install.BaseDirectory)
        }
    }

    foreach ($prefix in @('D:\Postgres17', 'C:\Postgres17')) {
        if (Test-Path $prefix) {
            $preferredPrefixes.Add($prefix)
        }
    }

    $psqlPath = Resolve-ExecutablePath -Name 'psql.exe' -PreferredPrefixes $preferredPrefixes.ToArray()
    $pgIsReadyPath = Resolve-ExecutablePath -Name 'pg_isready.exe' -PreferredPrefixes $preferredPrefixes.ToArray()
    $pgCtlPath = Resolve-ExecutablePath -Name 'pg_ctl.exe' -PreferredPrefixes $preferredPrefixes.ToArray()

    $matchedInstall = $null
    foreach ($install in $registryInstalls) {
        if ($psqlPath -and $install.BaseDirectory -and $psqlPath.StartsWith($install.BaseDirectory, [System.StringComparison]::OrdinalIgnoreCase)) {
            $matchedInstall = $install
            break
        }
    }

    $prefix = if ($psqlPath) { Split-Path -Parent (Split-Path -Parent $psqlPath) } else { $null }
    $serviceName = if ($matchedInstall -and $matchedInstall.Service) {
        $matchedInstall.Service
    }
    elseif ($prefix -and $prefix -match '[CD]:\\Postgres17$') {
        'postgresql-17'
    }
    else {
        $null
    }

    return @{
        PsqlPath = $psqlPath
        PgIsReadyPath = $pgIsReadyPath
        PgCtlPath = $pgCtlPath
        Prefix = $prefix
        RegistryInstall = $matchedInstall
        ServiceName = $serviceName
    }
}

function Ensure-PostgresServiceRunning {
    param([Parameter(Mandatory = $true)][hashtable]$Tooling)

    if ([string]::IsNullOrWhiteSpace($Tooling.ServiceName)) {
        return
    }

    $service = Get-Service -Name $Tooling.ServiceName -ErrorAction SilentlyContinue
    if (-not $service) {
        Write-Warn "Configured PostgreSQL service was not found: $($Tooling.ServiceName)"
        return
    }

    if ($service.Status -eq 'Running') {
        Write-OK "PostgreSQL service running: $($Tooling.ServiceName)"
        return
    }

    if (-not (Test-IsAdministrator)) {
        Restart-CurrentProcessElevated -Reason "starting PostgreSQL service '$($Tooling.ServiceName)'"
    }

    Write-Info "Starting PostgreSQL service: $($Tooling.ServiceName)"
    Start-Service -Name $Tooling.ServiceName -ErrorAction Stop

    $service.WaitForStatus('Running', [TimeSpan]::FromSeconds(20))
    Write-OK "PostgreSQL service running: $($Tooling.ServiceName)"
}

function Read-InteractiveValue {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [string]$Default = $null,
        [switch]$Secret
    )

    if (-not [Environment]::UserInteractive) {
        if ($Default) {
            return $Default
        }

        throw "Interactive input is required for: $Prompt"
    }

    $fullPrompt = if ([string]::IsNullOrWhiteSpace($Default)) { $Prompt } else { "$Prompt [$Default]" }

    try {
        if ($Secret) {
            $secure = Read-Host $Prompt -AsSecureString
            if ($secure) {
                $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
                try {
                    return [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
                }
                finally {
                    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
                }
            }
        }
        else {
            $value = Read-Host $fullPrompt
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value.Trim()
            }
            return $Default
        }
    }
    catch {
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'CNPJ ETL Pipeline'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(440, 170)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $fullPrompt
    $label.Location = New-Object System.Drawing.Point(12, 15)
    $label.Size = New-Object System.Drawing.Size(400, 32)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(15, 55)
    $textBox.Size = New-Object System.Drawing.Size(395, 24)
    if ($Default) {
        $textBox.Text = $Default
        $textBox.SelectAll()
    }
    if ($Secret) {
        $textBox.UseSystemPasswordChar = $true
    }

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = 'OK'
    $okButton.Location = New-Object System.Drawing.Point(254, 95)
    $okButton.Size = New-Object System.Drawing.Size(75, 26)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = 'Cancel'
    $cancelButton.Location = New-Object System.Drawing.Point(335, 95)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 26)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $form.Controls.Add($label)
    $form.Controls.Add($textBox)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "Input cancelled for: $Prompt"
    }

    $value = $textBox.Text
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $Default
    }

    return $value.Trim()
}

function Get-PostgresPortCandidates {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [Parameter(Mandatory = $true)][hashtable]$Tooling
    )

    $ports = New-Object System.Collections.Generic.List[string]
    $seen = @{}

    function Add-PortCandidate {
        param($Value)

        if ($null -eq $Value) {
            return
        }

        $port = ([string]$Value).Trim()
        if ($port -notmatch '^\d+$') {
            return
        }

        if ($seen.ContainsKey($port)) {
            return
        }

        $seen[$port] = $true
        $ports.Add($port)
    }

    Add-PortCandidate -Value (Get-EnvConfigValue -Names @('CNPJ_ETL_DB_PORT', 'PGPORT'))
    if ($Config.ContainsKey('dbPort')) {
        Add-PortCandidate -Value $Config['dbPort']
    }

    if ($Tooling.RegistryInstall -and $Tooling.RegistryInstall.Port) {
        Add-PortCandidate -Value $Tooling.RegistryInstall.Port
    }

    if ($Tooling.Prefix -and $Tooling.Prefix -match '[CD]:\\Postgres17$') {
        Add-PortCandidate -Value '5253'
    }

    $postgresProcesses = Get-Process postgres -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
    foreach ($processId in $postgresProcesses) {
        foreach ($conn in (Get-NetTCPConnection -State Listen -OwningProcess $processId -ErrorAction SilentlyContinue)) {
            Add-PortCandidate -Value $conn.LocalPort
        }
    }

    Add-PortCandidate -Value '5253'
    Add-PortCandidate -Value '5432'

    return $ports.ToArray()
}

function Resolve-DatabaseRuntimeConfig {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [Parameter(Mandatory = $true)][hashtable]$Tooling
    )

    $resolvedHost = Get-EnvConfigValue -Names @('CNPJ_ETL_DB_HOST', 'PGHOST') -Default $Config['dbHost']
    if ($Tooling.Prefix -and $Tooling.Prefix -match '[CD]:\\Postgres17$' -and ($resolvedHost -eq 'localhost' -or [string]::IsNullOrWhiteSpace($resolvedHost))) {
        $resolvedHost = '127.0.0.1'
    }
    elseif ([string]::IsNullOrWhiteSpace($resolvedHost)) {
        $resolvedHost = '127.0.0.1'
    }
    $Config['dbHost'] = $resolvedHost

    $defaultDbUser = if ($Config.ContainsKey('dbUser')) { $Config['dbUser'] } else { $null }
    $resolvedUser = Get-EnvConfigValue -Names @('CNPJ_ETL_DB_USER', 'PGUSER') -Default $defaultDbUser
    if ([string]::IsNullOrWhiteSpace($resolvedUser)) {
        $resolvedUser = Read-InteractiveValue -Prompt 'Database user' -Default 'postgres'
    }
    $Config['dbUser'] = $resolvedUser

    $defaultDbPassword = if ($Config.ContainsKey('dbPassword')) { $Config['dbPassword'] } else { $null }
    $resolvedPassword = Get-EnvConfigValue -Names @('CNPJ_ETL_DB_PASSWORD', 'PGPASSWORD') -Default $defaultDbPassword
    if ([string]::IsNullOrWhiteSpace($resolvedPassword)) {
        $resolvedPassword = Read-InteractiveValue -Prompt "Database password for $resolvedUser" -Secret
    }
    $Config['dbPassword'] = $resolvedPassword

    $portCandidates = Get-PostgresPortCandidates -Config $Config -Tooling $Tooling
    if ($portCandidates.Count -eq 0) {
        $portCandidates = @('5253', '5432')
    }

    $errors = New-Object System.Collections.Generic.List[string]
    foreach ($port in $portCandidates) {
        $Config['dbPort'] = $port
        try {
            Test-DatabaseConnection -PsqlPath $Tooling.PsqlPath -Config $Config -Database 'postgres' -Query 'SELECT version();' -ExpectedPattern 'PostgreSQL'
            Write-Info "PostgreSQL connection selected on port $port"
            return
        }
        catch {
            $errors.Add("${port}: $($_.Exception.Message)")
        }
    }

    throw "Unable to connect to PostgreSQL using the detected settings. Tried ports: $($portCandidates -join ', '). Details: $($errors -join ' | ')"
}

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-DriveInfoForPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $root = Split-Path -Path $Path -Qualifier
    if (-not $root) {
        $root = [System.IO.Path]::GetPathRoot($Path)
    }

    if (-not $root -or -not (Test-Path $root)) {
        throw "Drive not found for path: $Path"
    }

    return [System.IO.DriveInfo]::new($root)
}

function Assert-DirectoryWritable {
    param([Parameter(Mandatory = $true)][string]$Path)

    Ensure-Directory -Path $Path
    $probePath = Join-Path $Path ".write_test_$PID.tmp"
    try {
        [System.IO.File]::WriteAllText($probePath, 'ok')
        Write-OK "Writable directory: $Path"
    }
    catch {
        throw "No write permission for: $Path. $($_.Exception.Message)"
    }
    finally {
        Remove-Item $probePath -Force -ErrorAction SilentlyContinue
    }
}

function Assert-FreeSpace {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][double]$MinimumGB,
        [string]$Label = $Path
    )

    $drive = Get-DriveInfoForPath -Path $Path
    $freeGB = [math]::Round($drive.AvailableFreeSpace / 1GB, 2)
    if ($freeGB -lt $MinimumGB) {
        throw "Insufficient free space in $Label. Available: ${freeGB}GB, minimum: ${MinimumGB}GB."
    }

    Write-OK "Free space in ${Label}: ${freeGB}GB"
}

function Test-MemoryAvailability {
    param(
        [double]$WarnBelowGB = 4
    )

    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $freeGB = [math]::Round((($os.FreePhysicalMemory * 1KB) / 1GB), 2)
    $totalGB = [math]::Round((($os.TotalVisibleMemorySize * 1KB) / 1GB), 2)

    if ($freeGB -lt $WarnBelowGB) {
        Write-Warn "Available RAM is low: ${freeGB}GB of ${totalGB}GB. Processing may be slow."
        return
    }

    Write-OK "Available RAM: ${freeGB}GB of ${totalGB}GB"
}

function Set-PostgresPasswordEnv {
    param([Parameter(Mandatory = $true)][hashtable]$Config)

    if ($Config.ContainsKey('dbPassword') -and -not [string]::IsNullOrWhiteSpace([string]$Config.dbPassword)) {
        $env:PGPASSWORD = $Config.dbPassword
    }
    else {
        Remove-Item Env:PGPASSWORD -ErrorAction SilentlyContinue
    }
}

function Wait-PostgresReady {
    param(
        [string]$PgIsReadyPath,
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [int]$TimeoutSec = 15
    )

    if (-not $PgIsReadyPath -or -not (Test-Path $PgIsReadyPath)) {
        return $false
    }

    Set-PostgresPasswordEnv -Config $Config
    $args = @(
        '-h', $Config.dbHost,
        '-p', [string]$Config.dbPort,
        '-d', $Config.dbName,
        '-U', $Config.dbUser,
        '-q'
    )

    $elapsed = 0
    while ($elapsed -lt $TimeoutSec) {
        $job = Start-Job -ScriptBlock {
            param($exe, $argList)
            & $exe $argList 2>$null; $LASTEXITCODE
        } -ArgumentList $PgIsReadyPath, $args

        $completed = Wait-Job -Job $job -Timeout 5
        if ($completed) {
            $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            if ($result -eq 0) {
                return $true
            }
        } else {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }

        Start-Sleep -Seconds 1
        $elapsed++
    }

    return $false
}

function Test-DatabaseConnection {
    param(
        [Parameter(Mandatory = $true)][string]$PsqlPath,
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [string]$Database = $null,
        [string]$Query = 'SELECT 1;',
        [string]$ExpectedPattern = '^1$',
        [int]$TimeoutSec = 10
    )

    Set-PostgresPasswordEnv -Config $Config
    if (-not $Database) {
        $Database = $Config.dbName
    }

    $password = if ($Config.dbPassword) { $Config.dbPassword } else { '' }
    $args = @(
        '-h', $Config.dbHost,
        '-p', [string]$Config.dbPort,
        '-d', $Database,
        '-U', $Config.dbUser,
        '-w',
        '-tAc', $Query
    )

    try {
        $job = Start-Job -ScriptBlock {
            param($exe, $argList)
            $output = & $exe $argList 2>&1
            [pscustomobject]@{
                ExitCode = $LASTEXITCODE
                Output = ($output | Out-String)
            }
        } -ArgumentList $PsqlPath, $args

        $completed = Wait-Job -Job $job -Timeout $TimeoutSec
        if (-not $completed) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            throw "PostgreSQL connection timed out after ${TimeoutSec}s"
        }

        $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue

        $exitCode = 0
        $output = ''
        if ($result -and $result.PSObject.Properties['ExitCode']) {
            $exitCode = [int]$result.ExitCode
            $output = ([string]$result.Output).Trim()
        }
        else {
            $output = ($result | Out-String).Trim()
        }

        if ($exitCode -ne 0) {
            throw "PostgreSQL validation query failed with exit code $exitCode. Output: $output"
        }
        
        if ($output -notmatch $ExpectedPattern) {
            throw "PostgreSQL validation query returned unexpected result: $output"
        }

        Write-OK "PostgreSQL connectivity validated at $($Config.dbHost):$($Config.dbPort)/$Database"
    }
    catch {
        throw "Failed to connect to PostgreSQL: $($_.Exception.Message)"
    }
}

function Invoke-PsqlFileChecked {
    param(
        [Parameter(Mandatory = $true)][string]$PsqlPath,
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string]$Description = 'SQL file',
        [int]$TimeoutSec = 60
    )

    Set-PostgresPasswordEnv -Config $Config
    $args = @(
        '-h', $Config.dbHost,
        '-p', [string]$Config.dbPort,
        '-d', $Config.dbName,
        '-U', $Config.dbUser,
        '-w',
        '-f', $FilePath,
        '-v', 'ON_ERROR_STOP=1'
    )

    try {
        $job = Start-Job -ScriptBlock {
            param($exe, $argList)
            $output = & $exe $argList 2>&1
            [pscustomobject]@{
                ExitCode = $LASTEXITCODE
                Output = ($output | Out-String)
            }
        } -ArgumentList $PsqlPath, $args

        $completed = Wait-Job -Job $job -Timeout $TimeoutSec
        if (-not $completed) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            throw "SQL execution timed out after ${TimeoutSec}s"
        }

        $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue

        $exitCode = 0
        $jobOutput = ''
        if ($result -and $result.PSObject.Properties['ExitCode']) {
            $exitCode = [int]$result.ExitCode
            $jobOutput = [string]$result.Output
        }
        else {
            $jobOutput = ($result | Out-String)
        }

        $messageLines = @($jobOutput -split "`r?`n" | Where-Object { $_.Trim() -and $_ -notmatch '(^|:)\s*NOTICE\s*:' })
        $errorLines = @($messageLines | Where-Object { $_ -match '(^|:)\s*(ERROR|FATAL|PANIC)\s*:' })
        if ($exitCode -ne 0 -or $errorLines.Count -gt 0) {
            if ($messageLines.Count -gt 0) {
                throw "psql exited with code $exitCode. Output: $($messageLines -join [Environment]::NewLine)"
            }
            throw "psql exited with code $exitCode and no output."
        }

        if ($messageLines.Count -gt 0) {
            Write-Warn ($messageLines -join [Environment]::NewLine)
        }
    }
    catch {
        throw "Failed to execute $Description. $($_.Exception.Message)"
    }
}

function Invoke-CnpjPreflight {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [string[]]$EnsureDirectories = @(),
        [hashtable]$MinFreeSpaceGB = @{},
        [switch]$RequirePostgres,
        [switch]$RequireDatabase,
        [switch]$CheckMemory
    )

    Assert-ConfigValid -Config $Config

    foreach ($path in $EnsureDirectories) {
        Assert-DirectoryWritable -Path $path
    }

    foreach ($item in $MinFreeSpaceGB.GetEnumerator()) {
        Assert-FreeSpace -Path $item.Key -MinimumGB ([double]$item.Value) -Label $item.Key
    }

    if ($CheckMemory) {
        Test-MemoryAvailability
    }

    $tooling = @{}
    if ($RequirePostgres -or $RequireDatabase) {
        $tooling = Resolve-PostgresTooling -Config $Config
        $tooling.PsqlPath = $tooling.PsqlPath
        if (-not $tooling.PsqlPath) {
            throw 'psql.exe was not found in PATH, custom PostgreSQL installs, registry installs, or common PostgreSQL folders.'
        }
        Write-OK "psql found: $($tooling.PsqlPath)"

        if ($tooling.PgIsReadyPath) {
            Write-OK "pg_isready found: $($tooling.PgIsReadyPath)"
        }
        else {
            Write-Warn 'pg_isready.exe not found. Readiness checks will fall back to a direct psql connection.'
        }

        Ensure-PostgresServiceRunning -Tooling $tooling
        Resolve-DatabaseRuntimeConfig -Config $Config -Tooling $tooling
    }

    if ($RequireDatabase) {
        if ($tooling.PgIsReadyPath) {
            if (-not (Wait-PostgresReady -PgIsReadyPath $tooling.PgIsReadyPath -Config $Config)) {
                throw "PostgreSQL did not become ready in time at $($Config.dbHost):$($Config.dbPort)."
            }
            Write-OK 'PostgreSQL is accepting connections.'
        }

        Test-DatabaseConnection -PsqlPath $tooling.PsqlPath -Config $Config
    }

    return $tooling
}
