#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string[]]$ComputerName,
    [string]$Instance,
    [switch]$DiscoveryOnly,
    [switch]$EmitDiscoveryJson,
    [string]$TargetManifestBase64,
    [switch]$Force
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

function ConvertFrom-ODBCTargetManifest {
    param(
        [Parameter(Mandatory)][string]$Base64,
        [Parameter(Mandatory)][string[]]$AllowedServers
    )

    $json = [System.Text.Encoding]::UTF8.GetString(
        [System.Convert]::FromBase64String($Base64)
    )
    $decodedManifest = $json | ConvertFrom-Json -ErrorAction Stop
    $records = @($decodedManifest | ForEach-Object { $_ })
    if ($records.Count -eq 0) {
        throw 'The ODBC target manifest is empty.'
    }

    foreach ($record in $records) {
        $server = [string]$record.Server
        $architecture = [string]$record.Architecture
        $name = [string]$record.Name
        if ([string]::IsNullOrWhiteSpace($server) -or
            $architecture -notin @('32-bit', '64-bit') -or
            [string]::IsNullOrWhiteSpace($name)) {
            throw 'Every ODBC target must contain Server, Architecture, and Name.'
        }
        if ($AllowedServers -notcontains $server) {
            throw "ODBC target manifest server '$server' is not in ComputerName."
        }
    }

    return $records
}

function ConvertTo-ODBCRemoteTargetArgument {
    param([Parameter(Mandatory)][object[]]$Targets)

    $records = @(
        foreach ($target in @($Targets)) {
            $architecture = [string]$target.Architecture
            $name = [string]$target.Name
            if ($architecture -notin @('32-bit', '64-bit') -or
                [string]::IsNullOrWhiteSpace($name)) {
                throw 'Every remote ODBC target must contain one Architecture and one Name.'
            }

            [PSCustomObject]@{
                Architecture = $architecture
                Name         = $name
            }
        }
    )
    if ($records.Count -eq 0) {
        throw 'The remote ODBC target list is empty.'
    }

    $json = ConvertTo-Json -InputObject $records -Compress -Depth 3
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
}

function Confirm-DestructiveAction {
    param(
        [Parameter(Mandatory)][string]$Message,
        [switch]$Force
    )

    if ($Force) { return $true }
    while ($true) {
        $reply = Read-Host "$Message [y/N]"
        if ([string]::IsNullOrWhiteSpace($reply)) { return $false }
        switch ($reply.Trim().ToUpperInvariant()) {
            'Y' { return $true }
            'N' { return $false }
            default { Write-Host 'Please answer Y or N.' -ForegroundColor Yellow }
        }
    }
}

function Test-ODBCRemovalShouldProceed {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string]$Target,
        [Parameter(Mandatory)][string]$Action,
        [switch]$Force
    )

    # The GUI has already shown the exact DSN list and calls this worker with
    # -Force after the user confirms. Avoid a second PowerShell confirmation
    # prompt, which is unavailable in the GUI's non-interactive child process.
    # WhatIf must always win over Force so previews remain non-destructive.
    if ($Force -and -not $WhatIfPreference) { return $true }
    return $PSCmdlet.ShouldProcess($Target, $Action)
}

$discoveryScript = {
    param([string]$Token)

    $results = New-Object System.Collections.ArrayList
    foreach ($viewDefinition in @(
        [PSCustomObject]@{ Architecture = '64-bit'; View = [Microsoft.Win32.RegistryView]::Registry64 }
        [PSCustomObject]@{ Architecture = '32-bit'; View = [Microsoft.Win32.RegistryView]::Registry32 }
    )) {
        $baseKey = $null
        $sourcesKey = $null
        try {
            $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::LocalMachine,
                $viewDefinition.View
            )
            $sourcesKey = $baseKey.OpenSubKey('SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources', $false)
            if ($null -eq $sourcesKey) { continue }

            foreach ($dsnName in @($sourcesKey.GetValueNames() | Sort-Object)) {
                if ([string]::IsNullOrWhiteSpace($dsnName) -or
                    $dsnName.IndexOf($Token, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                    continue
                }

                $dsnKey = $null
                try {
                    $dsnKey = $baseKey.OpenSubKey(('SOFTWARE\ODBC\ODBC.INI\{0}' -f $dsnName), $false)
                    [void]$results.Add([PSCustomObject]@{
                        Architecture = [string]$viewDefinition.Architecture
                        Name         = [string]$dsnName
                        Driver       = [string]$sourcesKey.GetValue($dsnName, '')
                        KeyDriver    = if ($null -ne $dsnKey) { [string]$dsnKey.GetValue('Driver', '') } else { '' }
                        DataServer   = if ($null -ne $dsnKey) { [string]$dsnKey.GetValue('Server', '') } else { '' }
                        Database     = if ($null -ne $dsnKey) { [string]$dsnKey.GetValue('Database', '') } else { '' }
                        Description  = if ($null -ne $dsnKey) { [string]$dsnKey.GetValue('Description', '') } else { '' }
                        RegistryPath = if ($viewDefinition.Architecture -eq '64-bit') {
                            'HKLM\SOFTWARE\ODBC\ODBC.INI\{0}' -f $dsnName
                        } else {
                            'HKLM\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\{0}' -f $dsnName
                        }
                    })
                } finally {
                    if ($null -ne $dsnKey) { $dsnKey.Dispose() }
                }
            }
        } finally {
            if ($null -ne $sourcesKey) { $sourcesKey.Dispose() }
            if ($null -ne $baseKey) { $baseKey.Dispose() }
        }
    }
    return @($results)
}

$removeScript = {
    param(
        [string]$Token,
        [string]$EncodedTargets
    )

    $targetJson = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($EncodedTargets))
    $decodedTargets = $targetJson | ConvertFrom-Json -ErrorAction Stop
    $targets = @($decodedTargets | ForEach-Object { $_ })
    if ($targets.Count -eq 0) {
        throw 'The remote ODBC target list is empty.'
    }

    $results = New-Object System.Collections.ArrayList
    foreach ($target in $targets) {
        $architecture = [string]$target.Architecture
        $dsnName = [string]$target.Name
        if ($architecture -notin @('32-bit', '64-bit') -or
            [string]::IsNullOrWhiteSpace($dsnName)) {
            throw 'A remote ODBC target did not contain one valid Architecture and one Name.'
        }
        $view = if ($architecture -eq '64-bit') {
            [Microsoft.Win32.RegistryView]::Registry64
        } else {
            [Microsoft.Win32.RegistryView]::Registry32
        }

        $baseKey = $null
        $sourcesKey = $null
        try {
            if ($dsnName.IndexOf($Token, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                throw "The live DSN name no longer contains instance token '$Token'."
            }

            $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::LocalMachine,
                $view
            )
            $sourcesKey = $baseKey.OpenSubKey('SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources', $true)
            if ($null -eq $sourcesKey) {
                [void]$results.Add([PSCustomObject]@{
                    Architecture = $architecture; Name = $dsnName; Success = $true
                    Message = 'ODBC Data Sources registry key is absent; target was already removed.'
                })
                continue
            }

            $liveName = @($sourcesKey.GetValueNames() | Where-Object {
                [System.StringComparer]::OrdinalIgnoreCase.Equals([string]$_, $dsnName)
            } | Select-Object -First 1)
            if ($liveName.Count -eq 0) {
                [void]$results.Add([PSCustomObject]@{
                    Architecture = $architecture; Name = $dsnName; Success = $true
                    Message = 'System DSN is already absent.'
                })
                continue
            }

            $baseKey.DeleteSubKeyTree(('SOFTWARE\ODBC\ODBC.INI\{0}' -f [string]$liveName[0]), $false)
            $sourcesKey.DeleteValue([string]$liveName[0], $false)

            $dsnStillListed = @($sourcesKey.GetValueNames() | Where-Object {
                [System.StringComparer]::OrdinalIgnoreCase.Equals([string]$_, $dsnName)
            }).Count -gt 0
            $verificationKey = $baseKey.OpenSubKey(('SOFTWARE\ODBC\ODBC.INI\{0}' -f $dsnName), $false)
            $dsnKeyStillExists = $null -ne $verificationKey
            if ($null -ne $verificationKey) { $verificationKey.Dispose() }
            if ($dsnStillListed -or $dsnKeyStillExists) {
                throw 'Post-removal verification found the DSN registry entry still present.'
            }

            [void]$results.Add([PSCustomObject]@{
                Architecture = $architecture; Name = $dsnName; Success = $true
                Message = 'System DSN registry key and ODBC Data Sources entry were removed.'
            })
        } catch {
            [void]$results.Add([PSCustomObject]@{
                Architecture = $architecture; Name = $dsnName; Success = $false
                Message = $_.Exception.Message
            })
        } finally {
            if ($null -ne $sourcesKey) { $sourcesKey.Dispose() }
            if ($null -ne $baseKey) { $baseKey.Dispose() }
        }
    }
    return @($results)
}

$servers = @(ConvertTo-ServerList -InputNames $ComputerName)
if ($servers.Count -eq 0) {
    Write-Error 'No servers provided.'
    exit 1
}
foreach ($server in $servers) {
    if ($server -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        Write-Error ("Invalid server name '{0}'." -f $server)
        exit 1
    }
}

if ([string]::IsNullOrWhiteSpace($Instance)) {
    $Instance = Read-Host 'Enter the instance token to match in System DSN names'
}
$Instance = $Instance.Trim()
if ($Instance.Length -lt 3 -or $Instance -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
    Write-Error 'The instance token must be at least 3 characters and contain only letters, numbers, period, underscore, or hyphen.'
    exit 1
}
if ($EmitDiscoveryJson -and -not $DiscoveryOnly) {
    Write-Error '-EmitDiscoveryJson can be used only with -DiscoveryOnly.'
    exit 1
}

$manifest = @()
$hasManifest = -not [string]::IsNullOrWhiteSpace($TargetManifestBase64)
if ($hasManifest) {
    try {
        $manifest = @(ConvertFrom-ODBCTargetManifest -Base64 $TargetManifestBase64 -AllowedServers $servers)
    } catch {
        Write-Error ("Invalid ODBC target manifest: {0}" -f $_.Exception.Message)
        exit 1
    }
}

Write-Host ("Servers: {0}" -f ($servers -join ', '))
Write-Host ("System DSN instance token: {0}" -f $Instance)
if ($WhatIfPreference) {
    Write-Host '[WhatIf] Preview mode is ON. No 32-bit or 64-bit System DSN will be removed.'
}

$hadFailures = $false
foreach ($server in $servers) {
    Write-Host ("==== {0} :: ODBC System DSNs ====" -f $server) -ForegroundColor Cyan
    try {
        $liveTargets = @(Invoke-Command -ComputerName $server -ScriptBlock $discoveryScript -ArgumentList $Instance -ErrorAction Stop)
    } catch {
        Write-Warning ("[{0}] ODBC System DSN discovery failed: {1}" -f $server, $_.Exception.Message)
        $hadFailures = $true
        continue
    }

    if ($hasManifest) {
        $serverManifest = @($manifest | Where-Object { [string]$_.Server -eq $server })
        $liveTargets = @(
            foreach ($requested in $serverManifest) {
                $match = @($liveTargets | Where-Object {
                    [string]$_.Architecture -eq [string]$requested.Architecture -and
                    [System.StringComparer]::OrdinalIgnoreCase.Equals([string]$_.Name, [string]$requested.Name)
                } | Select-Object -First 1)
                if ($match.Count -gt 0) {
                    $match[0]
                } else {
                    Write-Host ("[{0}] Confirmed {1} System DSN '{2}' is absent or no longer matches; skipped." -f $server, $requested.Architecture, $requested.Name)
                }
            }
        )
    }

    if ($liveTargets.Count -eq 0) {
        Write-Host ("[{0}] No 32-bit or 64-bit System DSNs matched '{1}'." -f $server, $Instance)
        continue
    }

    foreach ($target in $liveTargets) {
        Write-Host ("  {0} | {1} | Driver={2} | Server={3} | Database={4}" -f $target.Architecture, $target.Name, $target.Driver, $target.DataServer, $target.Database)
        if ($DiscoveryOnly -and $EmitDiscoveryJson) {
            $record = [PSCustomObject]@{
                Server       = $server
                Architecture = [string]$target.Architecture
                Name         = [string]$target.Name
                Driver       = [string]$target.Driver
                KeyDriver    = [string]$target.KeyDriver
                DataServer   = [string]$target.DataServer
                Database     = [string]$target.Database
                Description  = [string]$target.Description
                RegistryPath = [string]$target.RegistryPath
            }
            Write-Output ('@@ODBC_TARGET@@' + ($record | ConvertTo-Json -Compress))
        }
    }

    if ($DiscoveryOnly) { continue }

    $approvedTargets = New-Object System.Collections.ArrayList
    foreach ($target in $liveTargets) {
        $registryTarget = '{0}\{1}' -f $server, $target.RegistryPath
        $action = "Remove $($target.Architecture) System DSN '$($target.Name)'"
        if (-not (Test-ODBCRemovalShouldProceed -Target $registryTarget -Action $action -Force:$Force)) { continue }
        if (-not (Confirm-DestructiveAction -Message ("Remove {0} System DSN '{1}' from {2}?" -f $target.Architecture, $target.Name, $server) -Force:$Force)) {
            Write-Host ("Skipped System DSN '{0}'." -f $target.Name)
            continue
        }
        [void]$approvedTargets.Add([PSCustomObject]@{
            Architecture = [string]$target.Architecture
            Name = [string]$target.Name
        })
    }

    if ($approvedTargets.Count -eq 0) { continue }
    try {
        # Pass one scalar encoded manifest so PowerShell remoting cannot merge
        # same-named 32-bit and 64-bit DSNs into array-valued properties.
        $encodedApprovedTargets = ConvertTo-ODBCRemoteTargetArgument -Targets @($approvedTargets)
        $removeResults = @(Invoke-Command -ComputerName $server -ScriptBlock $removeScript -ArgumentList $Instance, $encodedApprovedTargets -ErrorAction Stop)
    } catch {
        Write-Warning ("[{0}] ODBC System DSN removal failed: {1}" -f $server, $_.Exception.Message)
        $hadFailures = $true
        continue
    }

    foreach ($result in $removeResults) {
        if ($result.Success) {
            Write-Host ("[{0}] {1} System DSN '{2}': {3}" -f $server, $result.Architecture, $result.Name, $result.Message) -ForegroundColor Green
        } else {
            Write-Warning ("[{0}] Failed to remove {1} System DSN '{2}': {3}" -f $server, $result.Architecture, $result.Name, $result.Message)
            $hadFailures = $true
        }
    }
}

if ($hadFailures) {
    Write-Warning 'ODBC System DSN operation completed with one or more failures.'
    exit 2
}
if ($DiscoveryOnly) {
    Write-Host 'ODBC System DSN discovery completed successfully.'
} elseif ($WhatIfPreference) {
    Write-Host 'ODBC System DSN preview completed successfully.'
} else {
    Write-Host 'ODBC System DSN cleanup completed successfully.'
}
