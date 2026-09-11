#Requires -Version 5.1

[CmdletBinding()]
param(
    [string[]]$ServiceComputerName,
    [string[]]$ApplicationComputerName,
    [Parameter(Mandatory)][string]$Instance,
    [ValidateSet('Token', 'Exact', 'Contains')][string]$IISMatchMode = 'Token'
)

$ErrorActionPreference = 'Stop'

function ConvertTo-ServerList {
    param([string[]]$InputNames)

    @(
        $InputNames |
            ForEach-Object { $_ -split '[,;\s]+' } |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ } |
            Select-Object -Unique
    )
}

function Write-ScanRecord {
    param([Parameter(Mandatory)]$Record)

    Write-Output ('@@SCAN_TARGET@@' + ($Record | ConvertTo-Json -Compress -Depth 5))
}

function Invoke-StructuredDiscovery {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string[]]$Arguments,
        [Parameter(Mandatory)][string]$Marker,
        [Parameter(Mandatory)][scriptblock]$ConvertRecord
    )

    Write-Output ("=== Scan - {0} ===" -f $Title)
    $powerShellExe = Join-Path -Path $PSHOME -ChildPath 'powershell.exe'
    $processArguments = @(
        '-NoLogo', '-NoProfile', '-NonInteractive',
        '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath
    ) + $Arguments

    try {
        & $powerShellExe @processArguments 2>&1 | ForEach-Object {
            $line = [string]$_
            if ($line.StartsWith($Marker, [System.StringComparison]::Ordinal)) {
                try {
                    $source = $line.Substring($Marker.Length) | ConvertFrom-Json -ErrorAction Stop
                    $convertedRecord = & $ConvertRecord $source
                    if ($null -ne $convertedRecord) {
                        Write-ScanRecord -Record $convertedRecord
                    }
                } catch {
                    Write-Output ("ERROR: Invalid {0} discovery record: {1}" -f $Title, $_.Exception.Message)
                    $script:HadFailures = $true
                }
            } else {
                Write-Output $line
            }
        }
        if ($LASTEXITCODE -ne 0) {
            Write-Output ("ERROR: {0} discovery finished with exit code {1}" -f $Title, $LASTEXITCODE)
            $script:HadFailures = $true
        }
    } catch {
        Write-Output ("ERROR: Unable to run {0} discovery: {1}" -f $Title, $_.Exception.Message)
        $script:HadFailures = $true
    }
}

$serviceServers = @(ConvertTo-ServerList -InputNames $ServiceComputerName)
$applicationServers = @(ConvertTo-ServerList -InputNames $ApplicationComputerName)
$combinedServers = @($serviceServers + $applicationServers | Select-Object -Unique)

if ($combinedServers.Count -eq 0) {
    Write-Error 'Enter at least one service or application server.'
    exit 1
}

$Instance = $Instance.Trim()
if ($Instance.Length -lt 3 -or $Instance -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
    Write-Error 'The instance token must be at least 3 characters and contain only letters, numbers, period, underscore, or hyphen.'
    exit 1
}

foreach ($server in $combinedServers) {
    if ($server -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        Write-Error ("Invalid server name '{0}'." -f $server)
        exit 1
    }
}

$serviceScript = Join-Path $PSScriptRoot 'Service deletion.ps1'
$iisScript = Join-Path $PSScriptRoot 'IIS site and pool deletion.ps1'
$localAdminScript = Join-Path $PSScriptRoot 'Local Admin Cleanup.ps1'
$odbcScript = Join-Path $PSScriptRoot 'ODBC System DSN Cleanup.ps1'
foreach ($requiredPath in @($serviceScript, $iisScript, $localAdminScript, $odbcScript)) {
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        Write-Error "Required scan dependency is missing: $requiredPath"
        exit 1
    }
}

$script:HadFailures = $false

if ($serviceServers.Count -gt 0) {
    Invoke-StructuredDiscovery `
        -Title 'Services' `
        -ScriptPath $serviceScript `
        -Arguments @(
            '-ComputerName', ($serviceServers -join ','),
            '-LogOnAsToken', $Instance,
            '-DiscoveryOnly', '-EmitDiscoveryJson', '-IncludeNameOnlyCandidates'
        ) `
        -Marker '@@SERVICE_TARGET@@' `
        -ConvertRecord {
            param($source)
            $serviceTargetType = if ([string]$source.TargetType -eq 'GCSListEntry') {
                'GCS list entry'
            } else {
                [string]$source.TargetType
            }
            [PSCustomObject]@{
                Category      = 'Services'
                Server        = [string]$source.Server
                TargetType    = $serviceTargetType
                Name          = [string]$source.Name
                DisplayName   = [string]$source.DisplayName
                State         = [string]$source.State
                StartName     = [string]$source.StartName
                FileName      = [string]$source.FileName
                Path          = [string]$source.Path
                Entry         = [string]$source.Entry
                SiteName      = ''
                AppPoolName   = ''
                Details       = ''
                BlockedReason = ''
                Planned       = $true
                MatchesToken  = $true
                MatchSource   = [string]$source.MatchSource
                MatchReason   = [string]$source.MatchReason
                IsOptInCandidate = [bool]$source.IsOptInCandidate
                OptInSelected = $false
                PoolIdentity  = ''
                IdentityMatchesToken = $false
            }
        }
}

if ($applicationServers.Count -gt 0) {
    Invoke-StructuredDiscovery `
        -Title 'IIS' `
        -ScriptPath $iisScript `
        -Arguments @(
            '-ComputerName', ($applicationServers -join ','),
            '-Instance', $Instance,
            '-MatchMode', $IISMatchMode,
            '-DiscoveryOnly', '-EmitDiscoveryJson'
        ) `
        -Marker '@@IIS_TARGET@@' `
        -ConvertRecord {
            param($source)
            $category = switch ([string]$source.TargetType) {
                'Site'        { 'IISSites' }
                'Application' { 'IISApplications' }
                'AppPool'     { 'ApplicationPools' }
                default       { 'IISSites' }
            }
            [PSCustomObject]@{
                Category      = $category
                Server        = [string]$source.Server
                TargetType    = [string]$source.TargetType
                Name          = [string]$source.Name
                DisplayName   = ''
                State         = [string]$source.State
                StartName     = ''
                FileName      = ''
                Path          = [string]$source.ApplicationPath
                Entry         = ''
                SiteName      = [string]$source.SiteName
                ApplicationPath = [string]$source.ApplicationPath
                AppPoolName   = [string]$source.AppPoolName
                Details       = [string]$source.Details
                BlockedReason = [string]$source.BlockedReason
                Planned       = [bool]$source.Planned
                MatchesToken  = $true
                MatchSource   = [string]$source.MatchSource
                MatchReason   = [string]$source.MatchReason
                IsOptInCandidate = [bool]$source.IsOptInCandidate
                OptInSelected = $false
                PoolIdentity  = [string]$source.PoolIdentity
                IdentityMatchesToken = [bool]$source.IdentityMatchesToken
            }
        }
}

Invoke-StructuredDiscovery `
    -Title 'Local administrators' `
    -ScriptPath $localAdminScript `
    -Arguments @(
        '-ComputerName', ($combinedServers -join ','),
        '-Instance', $Instance,
        '-DiscoveryOnly', '-EmitDiscoveryJson'
    ) `
    -Marker '@@LOCAL_ADMIN_TARGET@@' `
    -ConvertRecord {
        param($source)
        [PSCustomObject]@{
            Category      = 'LocalAdmins'
            Server        = [string]$source.Server
            TargetType    = [string]$source.Class
            Name          = [string]$source.Name
            DisplayName   = [string]$source.Group
            State         = ''
            StartName     = ''
            FileName      = ''
            Path          = [string]$source.Path
            Entry         = ''
            SiteName      = ''
            AppPoolName   = ''
            Details       = ''
            BlockedReason = ''
            Planned       = $true
            MatchesToken  = $true
        }
    }

Invoke-StructuredDiscovery `
    -Title 'ODBC System DSNs' `
    -ScriptPath $odbcScript `
    -Arguments @(
        '-ComputerName', ($combinedServers -join ','),
        '-Instance', $Instance,
        '-DiscoveryOnly', '-EmitDiscoveryJson'
    ) `
    -Marker '@@ODBC_TARGET@@' `
    -ConvertRecord {
        param($source)
        [PSCustomObject]@{
            Category      = 'ODBCDataSources'
            Server        = [string]$source.Server
            TargetType    = 'System DSN'
            Name          = [string]$source.Name
            DisplayName   = [string]$source.Driver
            State         = [string]$source.Architecture
            StartName     = [string]$source.DataServer
            FileName      = ''
            Path          = [string]$source.RegistryPath
            Entry         = ''
            SiteName      = ''
            AppPoolName   = ''
            Details       = [string]$source.Description
            BlockedReason = ''
            Planned       = $true
            MatchesToken  = $true
            Architecture  = [string]$source.Architecture
            Driver        = [string]$source.Driver
            DataServer    = [string]$source.DataServer
            Database      = [string]$source.Database
            RegistryPath  = [string]$source.RegistryPath
        }
    }

$remoteFolderScript = {
    $basePath = 'M:\OBOL\Clients'
    if (-not (Test-Path -LiteralPath $basePath -PathType Container -ErrorAction Stop)) {
        throw "Client folder root does not exist: $basePath"
    }
    @(Get-ChildItem -LiteralPath $basePath -Directory -Force -ErrorAction Stop | ForEach-Object { $_.Name } | Sort-Object -Unique)
}

Write-Output '=== Scan - Client folders ==='
foreach ($server in $combinedServers) {
    $names = @()
    try {
        $names = @(Invoke-Command -ComputerName $server -ScriptBlock $remoteFolderScript -ErrorAction Stop)
    } catch {
        Write-Output ("WARNING: Remote folder scan failed on {0}: {1}. Trying UNC." -f $server, $_.Exception.Message)
        $uncRoot = '\\{0}\m$\OBOL\Clients' -f $server
        try {
            if (-not (Test-Path -LiteralPath $uncRoot -PathType Container -ErrorAction Stop)) {
                throw "Client folder root does not exist: $uncRoot"
            }
            $names = @(Get-ChildItem -LiteralPath $uncRoot -Directory -Force -ErrorAction Stop | ForEach-Object { $_.Name } | Sort-Object -Unique)
        } catch {
            Write-Output ("ERROR: Unable to scan client folders on {0}: {1}" -f $server, $_.Exception.Message)
            $script:HadFailures = $true
            continue
        }
    }

    foreach ($name in @($names | ForEach-Object { [string]$_ } | Where-Object { $_ } | Sort-Object -Unique)) {
        $uncPath = '\\{0}\m$\OBOL\Clients\{1}' -f $server, $name
        Write-ScanRecord -Record ([PSCustomObject]@{
            Category      = 'Folders'
            Server        = [string]$server
            TargetType    = 'Folder'
            Name          = $name
            DisplayName   = ''
            State         = ''
            StartName     = ''
            FileName      = ''
            Path          = $uncPath
            Entry         = ''
            SiteName      = ''
            AppPoolName   = ''
            Details       = ''
            BlockedReason = ''
            Planned       = $false
            MatchesToken  = ($name.IndexOf($Instance, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
        })
    }
}

if ($script:HadFailures) {
    Write-Warning 'Scan completed with one or more discovery failures. Available results were retained.'
    exit 2
}

Write-Host 'Scan completed successfully. No configuration was changed.'
