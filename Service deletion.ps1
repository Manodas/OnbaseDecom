#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string[]]$ComputerName,

    [string]$LogOnAs,

    [string]$LogOnAsToken,

    [ValidateRange(1, 600)]
    [int]$StopTimeoutSec = 25,

    [switch]$AllowBuiltInAccount,

    [switch]$DiscoveryOnly,

    [switch]$EmitDiscoveryJson,

    [switch]$IncludeNameOnlyCandidates,

    [string]$TargetManifestBase64,

    [switch]$Force
)

$script:AccountIdentityCache = @{}

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

function ConvertFrom-ServiceTargetManifest {
    param(
        [Parameter(Mandatory)][string]$Base64,
        [Parameter(Mandatory)][string[]]$AllowedServers
    )

    $manifestJson = [System.Text.Encoding]::UTF8.GetString(
        [System.Convert]::FromBase64String($Base64)
    )
    $decodedManifest = $manifestJson | ConvertFrom-Json -ErrorAction Stop

    # Windows PowerShell 5.1 can return a JSON array as one Object[] pipeline
    # object. Explicitly enumerate it so each target remains a separate record.
    $records = @(
        $decodedManifest |
            ForEach-Object { $_ }
    )
    if ($records.Count -eq 0) {
        throw 'The target manifest is empty.'
    }

    foreach ($targetRecord in $records) {
        $targetServer = [string]$targetRecord.Server
        $targetType = [string]$targetRecord.TargetType
        if ([string]::IsNullOrWhiteSpace($targetType)) {
            $targetType = 'Service'
            $targetRecord | Add-Member -NotePropertyName TargetType -NotePropertyValue $targetType -Force
        }
        if (
            [string]::IsNullOrWhiteSpace($targetServer) -or
            $targetType -notin @('Service', 'GCSListEntry')
        ) {
            throw 'Every target manifest record must contain Server and a valid TargetType.'
        }
        if ($AllowedServers -notcontains $targetServer) {
            throw "Target manifest server '$targetServer' is not in ComputerName."
        }
        if ($targetType -eq 'Service' -and
            [string]::IsNullOrWhiteSpace([string]$targetRecord.Name)) {
            throw 'Every Service target must contain Name.'
        }
        if ($targetType -eq 'Service') {
            $matchSource = [string]$targetRecord.MatchSource
            if ([string]::IsNullOrWhiteSpace($matchSource)) {
                $matchSource = 'Account'
                $targetRecord | Add-Member -NotePropertyName MatchSource -NotePropertyValue $matchSource -Force
            }
            if ($matchSource -notin @('Account', 'NameOnly')) {
                throw "Service target '$([string]$targetRecord.Name)' has an invalid MatchSource."
            }
        }
        if ($targetType -eq 'GCSListEntry') {
            if ([string]$targetRecord.FileName -notin @('MonitoredServices.txt', 'ServicestoRestart.txt')) {
                throw 'Every GCSListEntry target must use an approved GCS list file.'
            }
            if ([string]::IsNullOrWhiteSpace([string]$targetRecord.Entry)) {
                throw 'Every GCSListEntry target must contain Entry.'
            }
        }
    }

    return $records
}

function Test-GCSServiceListEntryContainsToken {
    param(
        [string]$Entry,
        [string]$Token
    )

    if ([string]::IsNullOrWhiteSpace($Entry) -or
        [string]::IsNullOrWhiteSpace($Token)) {
        return $false
    }

    $trimmed = $Entry.Trim()
    if ($trimmed.StartsWith('#') -or $trimmed.StartsWith(';')) {
        return $false
    }

    return $trimmed.IndexOf(
        $Token.Trim(),
        [System.StringComparison]::OrdinalIgnoreCase
    ) -ge 0
}

function Get-GCSServiceListTargets {
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$Token
    )

    $root = '\\{0}\C$\OBOL\Utilities\GCSserviceStartRestarter' -f $Server
    foreach ($fileName in @('MonitoredServices.txt', 'ServicestoRestart.txt')) {
        $path = Join-Path -Path $root -ChildPath $fileName
        try {
            if (-not (Test-Path -LiteralPath $path -PathType Leaf -ErrorAction Stop)) {
                [PSCustomObject]@{
                    Success  = $true
                    Missing  = $true
                    FileName = $fileName
                    Path     = $path
                    Entries  = @()
                    Error    = ''
                }
                continue
            }

            $entries = @(
                [System.IO.File]::ReadAllLines($path) |
                    Where-Object {
                        Test-GCSServiceListEntryContainsToken -Entry ([string]$_) -Token $Token
                    }
            )
            [PSCustomObject]@{
                Success  = $true
                Missing  = $false
                FileName = $fileName
                Path     = $path
                Entries  = $entries
                Error    = ''
            }
        } catch {
            [PSCustomObject]@{
                Success  = $false
                Missing  = $false
                FileName = $fileName
                Path     = $path
                Entries  = @()
                Error    = $_.Exception.Message
            }
        }
    }
}

function Remove-ConfirmedGCSServiceListEntries {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$ConfirmedEntries,
        [Parameter(Mandatory)][string]$Token
    )

    $tempPath = '{0}.{1}.tmp' -f $Path, ([guid]::NewGuid().ToString('N'))
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        $offset = 0
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
            $encoding = New-Object System.Text.UTF8Encoding($true)
            $offset = 3
        } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
            $encoding = New-Object System.Text.UnicodeEncoding($false, $true)
            $offset = 2
        } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
            $encoding = New-Object System.Text.UnicodeEncoding($true, $true)
            $offset = 2
        } else {
            $encoding = [System.Text.Encoding]::Default
        }

        $text = $encoding.GetString($bytes, $offset, $bytes.Length - $offset)
        $newlineMatch = [regex]::Match($text, "\r\n|\n|\r")
        $newline = if ($newlineMatch.Success) { $newlineMatch.Value } else { [Environment]::NewLine }
        $lines = @([regex]::Split($text, "\r\n|\n|\r"))
        $confirmed = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
        foreach ($entry in $ConfirmedEntries) {
            [void]$confirmed.Add([string]$entry)
        }

        $removed = New-Object System.Collections.ArrayList
        $remaining = New-Object System.Collections.ArrayList
        foreach ($line in $lines) {
            if ($confirmed.Contains([string]$line) -and
                (Test-GCSServiceListEntryContainsToken -Entry ([string]$line) -Token $Token)) {
                [void]$removed.Add([string]$line)
            } else {
                [void]$remaining.Add([string]$line)
            }
        }

        if ($removed.Count -eq 0) {
            return [PSCustomObject]@{
                Success    = $true
                Removed    = @()
                BackupPath = ''
                Message    = 'No confirmed entries still matched; file was not changed.'
            }
        }

        $newText = @($remaining) -join $newline
        $preamble = $encoding.GetPreamble()
        $body = $encoding.GetBytes($newText)
        $newBytes = New-Object byte[] ($preamble.Length + $body.Length)
        [System.Array]::Copy($preamble, 0, $newBytes, 0, $preamble.Length)
        [System.Array]::Copy($body, 0, $newBytes, $preamble.Length, $body.Length)
        [System.IO.File]::WriteAllBytes($tempPath, $newBytes)

        $backupPath = '{0}.bak-{1}' -f $Path, (Get-Date -Format 'yyyyMMdd-HHmmssfff')
        [System.IO.File]::Replace($tempPath, $Path, $backupPath, $true)
        return [PSCustomObject]@{
            Success    = $true
            Removed    = @($removed)
            BackupPath = $backupPath
            Message    = 'Removed {0} confirmed entr{1}.' -f $removed.Count, $(if ($removed.Count -eq 1) { 'y' } else { 'ies' })
        }
    } catch {
        return [PSCustomObject]@{
            Success    = $false
            Removed    = @()
            BackupPath = ''
            Message    = $_.Exception.Message
        }
    } finally {
        if (Test-Path -LiteralPath $tempPath -PathType Leaf) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
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

function Normalize-Account {
    param(
        [string]$Account,
        [string]$Server
    )

    if ([string]::IsNullOrWhiteSpace($Account)) { return '' }

    $normalized = $Account.Trim()
    if ($normalized.StartsWith('.\')) {
        $normalized = '{0}\{1}' -f $Server, $normalized.Substring(2)
    }

    $normalized = $normalized.ToLowerInvariant()
    switch ($normalized) {
        'localsystem'    { return 'nt authority\system' }
        'system'         { return 'nt authority\system' }
        'localservice'   { return 'nt authority\localservice' }
        'networkservice' { return 'nt authority\networkservice' }
        default          { return $normalized }
    }
}

function Get-ServiceAccountNameComponent {
    param([string]$Account)

    if ([string]::IsNullOrWhiteSpace($Account)) { return '' }

    $value = $Account.Trim()
    $slashIndex = $value.LastIndexOf('\')
    if ($slashIndex -ge 0 -and $slashIndex -lt ($value.Length - 1)) {
        return $value.Substring($slashIndex + 1)
    }

    $atIndex = $value.IndexOf('@')
    if ($atIndex -gt 0) {
        return $value.Substring(0, $atIndex)
    }

    return $value
}

function Test-ServiceAccountContainsToken {
    param(
        [string]$Account,
        [string]$Token
    )

    if ([string]::IsNullOrWhiteSpace($Account) -or
        [string]::IsNullOrWhiteSpace($Token)) {
        return $false
    }

    $accountName = Get-ServiceAccountNameComponent -Account $Account
    return $accountName.IndexOf(
        $Token.Trim(),
        [System.StringComparison]::OrdinalIgnoreCase
    ) -ge 0
}

function Get-AccountIdentityKeys {
    param(
        [Parameter(Mandatory)][string]$Account,
        [Parameter(Mandatory)][string]$Server
    )

    $normalized = Normalize-Account -Account $Account -Server $Server
    if ([string]::IsNullOrWhiteSpace($normalized)) { return @() }

    $cacheKey = '{0}|{1}' -f $Server.ToLowerInvariant(), $normalized
    if ($script:AccountIdentityCache.ContainsKey($cacheKey)) {
        return @($script:AccountIdentityCache[$cacheKey])
    }

    $keys = @("name:$normalized")
    try {
        $ntAccount = New-Object System.Security.Principal.NTAccount -ArgumentList $normalized
        $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier])
        if ($sid) {
            $keys += "sid:$($sid.Value.ToLowerInvariant())"
        }
    } catch {
        # Exact normalized-name matching remains available if SID resolution
        # is unavailable for an offline domain or a remote local account.
    }

    $keys = @($keys | Select-Object -Unique)
    $script:AccountIdentityCache[$cacheKey] = $keys
    return $keys
}

function Test-IsProtectedServiceAccount {
    param(
        [Parameter(Mandatory)][string]$Account,
        [Parameter(Mandatory)][string]$Server
    )

    $normalized = Normalize-Account -Account $Account -Server $Server
    if ($normalized -in @(
        'nt authority\system',
        'nt authority\localservice',
        'nt authority\networkservice'
    )) {
        return $true
    }

    $keys = @(Get-AccountIdentityKeys -Account $Account -Server $Server)
    foreach ($key in $keys) {
        if ($key -in @('sid:s-1-5-18', 'sid:s-1-5-19', 'sid:s-1-5-20')) {
            return $true
        }
    }

    return $false
}

function Get-ServicesWithCimProtocol {
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][ValidateSet('WsMan', 'Dcom')][string]$Protocol
    )

    $session = $null
    try {
        $option = New-CimSessionOption -Protocol $Protocol
        $session = New-CimSession -ComputerName $Server -SessionOption $option -ErrorAction Stop
        return @(Get-CimInstance -ClassName Win32_Service -CimSession $session -ErrorAction Stop)
    } finally {
        if ($null -ne $session) {
            $session | Remove-CimSession -ErrorAction SilentlyContinue
        }
    }
}

function Get-RemoteServices {
    param([Parameter(Mandatory)][string]$Server)

    $errors = @()
    foreach ($protocol in @('WsMan', 'Dcom')) {
        try {
            return @(Get-ServicesWithCimProtocol -Server $Server -Protocol $protocol)
        } catch {
            $errors += '{0}: {1}' -f $protocol, $_.Exception.Message
        }
    }

    throw "Unable to query services. $($errors -join ' | ')"
}

function Stop-RemoteServiceSafe {
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$ServiceName,
        [Parameter(Mandatory)][int]$TimeoutSec
    )

    try {
        $controller = Get-Service -ComputerName $Server -Name $ServiceName -ErrorAction Stop
        if ($controller.Status -ne [System.ServiceProcess.ServiceControllerStatus]::Stopped) {
            Stop-Service -InputObject $controller -Force -ErrorAction Stop
            $controller.WaitForStatus(
                [System.ServiceProcess.ServiceControllerStatus]::Stopped,
                [TimeSpan]::FromSeconds($TimeoutSec)
            )
            $controller.Refresh()
        }

        if ($controller.Status -ne [System.ServiceProcess.ServiceControllerStatus]::Stopped) {
            return [PSCustomObject]@{
                Success = $false
                Message = "Service did not stop. Current state: $($controller.Status)"
            }
        }

        return [PSCustomObject]@{
            Success = $true
            Message = 'Service is stopped.'
        }
    } catch {
        return [PSCustomObject]@{
            Success = $false
            Message = $_.Exception.Message
        }
    }
}

function Remove-RemoteService {
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$ServiceName
    )

    try {
        $deleteOutput = @(& sc.exe "\\$Server" delete $ServiceName 2>&1)
        $deleteCode = $LASTEXITCODE
        if ($deleteCode -ne 0) {
            return [PSCustomObject]@{
                Success = $false
                Message = "sc.exe delete failed with exit code $deleteCode. $($deleteOutput -join ' ')"
            }
        }

        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        do {
            Start-Sleep -Milliseconds 500
            & sc.exe "\\$Server" query $ServiceName 2>&1 | Out-Null
            $queryCode = $LASTEXITCODE
            if ($queryCode -eq 1060) {
                return [PSCustomObject]@{
                    Success = $true
                    Message = 'Service removed and absence verified.'
                }
            }
        } while ($stopwatch.Elapsed.TotalSeconds -lt 15)

        return [PSCustomObject]@{
            Success = $false
            Message = "Service deletion was accepted but the service remained visible after 15 seconds. Last query exit code: $queryCode"
        }
    } catch {
        return [PSCustomObject]@{
            Success = $false
            Message = $_.Exception.Message
        }
    }
}

if (-not $ComputerName -or $ComputerName.Count -eq 0) {
    $ComputerName = @((Read-Host 'Enter service servers (comma, space, or semicolon separated)'))
}
$servers = @(ConvertTo-ServerList -InputNames $ComputerName)
if ($servers.Count -eq 0) {
    Write-Error 'No service servers provided.'
    exit 1
}

foreach ($server in $servers) {
    if ($server -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        Write-Error ("Invalid server name '{0}'." -f $server)
        exit 1
    }
}

if (-not [string]::IsNullOrWhiteSpace($LogOnAs) -and
    -not [string]::IsNullOrWhiteSpace($LogOnAsToken)) {
    Write-Error 'Specify either -LogOnAs for an exact identity or -LogOnAsToken for an account-name token, not both.'
    exit 1
}

if ([string]::IsNullOrWhiteSpace($LogOnAs) -and
    [string]::IsNullOrWhiteSpace($LogOnAsToken)) {
    $LogOnAs = Read-Host "Enter the service 'Log On As' account"
}
if ([string]::IsNullOrWhiteSpace($LogOnAs) -and
    [string]::IsNullOrWhiteSpace($LogOnAsToken)) {
    Write-Error 'No service account provided.'
    exit 1
}

$useAccountToken = -not [string]::IsNullOrWhiteSpace($LogOnAsToken)
if ($useAccountToken) {
    $LogOnAsToken = $LogOnAsToken.Trim()
    if ($LogOnAsToken.Length -lt 3 -or $LogOnAsToken -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        Write-Error 'The Log On As token must be at least 3 characters and contain only letters, numbers, period, underscore, or hyphen.'
        exit 1
    }
} else {
    $LogOnAs = $LogOnAs.Trim()
}

if ($EmitDiscoveryJson -and -not $DiscoveryOnly) {
    Write-Error '-EmitDiscoveryJson can be used only with -DiscoveryOnly.'
    exit 1
}

$targetManifest = @()
$hasTargetManifest = -not [string]::IsNullOrWhiteSpace($TargetManifestBase64)
if ($hasTargetManifest) {
    try {
        $targetManifest = @(
            ConvertFrom-ServiceTargetManifest `
                -Base64 $TargetManifestBase64 `
                -AllowedServers $servers
        )
    } catch {
        Write-Error ("Invalid service target manifest: {0}" -f $_.Exception.Message)
        exit 1
    }
}

if (-not $useAccountToken -and -not $AllowBuiltInAccount) {
    foreach ($server in $servers) {
        if (Test-IsProtectedServiceAccount -Account $LogOnAs -Server $server) {
            Write-Error ("Refusing to target protected Windows service account '{0}'. Use -AllowBuiltInAccount only if this is intentional." -f $LogOnAs)
            exit 2
        }
    }
}

Write-Host ("Service servers: {0}" -f ($servers -join ', '))
if ($useAccountToken) {
    Write-Host ("Target Log On As account-name token: {0}" -f $LogOnAsToken)
} else {
    Write-Host ("Target Log On As: {0}" -f $LogOnAs)
}
if ($WhatIfPreference) {
    Write-Host '[WhatIf] Preview mode is ON. No service will be stopped or deleted, and no GCS service-list file will be changed.'
}

$hadFailures = $false

foreach ($server in $servers) {
    Write-Host ("==== {0} :: Services ====" -f $server) -ForegroundColor Cyan

    try {
        $services = @(Get-RemoteServices -Server $server)
    } catch {
        Write-Warning ("[{0}] {1}" -f $server, $_.Exception.Message)
        $hadFailures = $true
        continue
    }

    $targetKeys = if ($useAccountToken) {
        @()
    } else {
        @(Get-AccountIdentityKeys -Account $LogOnAs -Server $server)
    }
    $accountMatches = @(
        foreach ($service in $services) {
            if ([string]::IsNullOrWhiteSpace([string]$service.StartName)) { continue }

            if ($useAccountToken) {
                if (-not (Test-IsProtectedServiceAccount -Account ([string]$service.StartName) -Server $server) -and
                    (Test-ServiceAccountContainsToken -Account ([string]$service.StartName) -Token $LogOnAsToken)) {
                    $service
                }
                continue
            }

            $serviceKeys = @(Get-AccountIdentityKeys -Account ([string]$service.StartName) -Server $server)
            foreach ($key in $serviceKeys) {
                if ($targetKeys -contains $key) {
                    $service
                    break
                }
            }
        }
    )
    $nameOnlyMatches = @()
    if ($useAccountToken) {
        $nameOnlyMatches = @(
            foreach ($service in $services) {
                $serviceName = [string]$service.Name
                $accountMatchesToken = Test-ServiceAccountContainsToken `
                    -Account ([string]$service.StartName) `
                    -Token $LogOnAsToken
                if (
                    -not [string]::IsNullOrWhiteSpace($serviceName) -and
                    $serviceName.IndexOf($LogOnAsToken, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -and
                    -not $accountMatchesToken
                ) {
                    $service
                }
            }
        )
    }
    $matches = @($accountMatches)
    if ($DiscoveryOnly -and $IncludeNameOnlyCandidates) {
        $matches = @($accountMatches + $nameOnlyMatches)
    }

    $gcsListTargets = @()
    if ($useAccountToken) {
        foreach ($listResult in @(Get-GCSServiceListTargets -Server $server -Token $LogOnAsToken)) {
            if (-not $listResult.Success) {
                Write-Warning ("[{0}] Failed to read {1}: {2}" -f $server, $listResult.Path, $listResult.Error)
                $hadFailures = $true
                continue
            }
            if ($listResult.Missing) {
                Write-Host ("[{0}] GCS service list not found; skipped: {1}" -f $server, $listResult.Path) -ForegroundColor DarkYellow
                continue
            }
            foreach ($entry in @($listResult.Entries)) {
                $gcsListTargets += [PSCustomObject]@{
                    Server   = $server
                    FileName = [string]$listResult.FileName
                    Path     = [string]$listResult.Path
                    Entry    = [string]$entry
                }
            }
        }
    }

    if ($hasTargetManifest) {
        $confirmedAccountNames = @(
            $targetManifest |
                Where-Object {
                    ([string]$_.Server) -eq $server -and
                    ([string]$_.TargetType) -eq 'Service' -and
                    ([string]$_.MatchSource) -ne 'NameOnly'
                } |
                ForEach-Object { [string]$_.Name } |
                Select-Object -Unique
        )
        $confirmedNameOnlyNames = @(
            $targetManifest |
                Where-Object {
                    ([string]$_.Server) -eq $server -and
                    ([string]$_.TargetType) -eq 'Service' -and
                    ([string]$_.MatchSource) -eq 'NameOnly'
                } |
                ForEach-Object { [string]$_.Name } |
                Select-Object -Unique
        )
        $confirmedNames = @($confirmedAccountNames + $confirmedNameOnlyNames | Select-Object -Unique)
        $matches = @(
            @($accountMatches | Where-Object { $confirmedAccountNames -contains ([string]$_.Name) }) +
            @($nameOnlyMatches | Where-Object { $confirmedNameOnlyNames -contains ([string]$_.Name) })
        )

        $currentMatchNames = @($matches | ForEach-Object { [string]$_.Name })
        foreach ($confirmedName in $confirmedNames) {
            if ($currentMatchNames -notcontains $confirmedName) {
                Write-Host (
                    "[{0}] Confirmed service '{1}' is now absent or no longer matches its confirmed account/name-only rule; skipped." -f
                        $server,
                        $confirmedName
                )
            }
        }

        $confirmedListKeys = @(
            $targetManifest |
                Where-Object {
                    ([string]$_.Server) -eq $server -and
                    ([string]$_.TargetType) -eq 'GCSListEntry'
                } |
                ForEach-Object { '{0}|{1}' -f ([string]$_.FileName), ([string]$_.Entry) }
        )
        $gcsListTargets = @(
            $gcsListTargets |
                Where-Object {
                    $confirmedListKeys -contains ('{0}|{1}' -f $_.FileName, $_.Entry)
                }
        )
    }

    if ($matches.Count -eq 0 -and $gcsListTargets.Count -eq 0) {
        $matchText = if ($useAccountToken) {
            "account-name token '$LogOnAsToken'"
        } else {
            "account '$LogOnAs'"
        }
        Write-Host ("[{0}] No services found for {1}." -f $server, $matchText)
        continue
    }

    if ($EmitDiscoveryJson) {
        foreach ($service in $matches) {
            $isNameOnlyCandidate = @($nameOnlyMatches | Where-Object { ([string]$_.Name) -eq ([string]$service.Name) }).Count -gt 0
            $discoveryRecord = [PSCustomObject]@{
                TargetType  = 'Service'
                Server      = $server
                Name        = [string]$service.Name
                DisplayName = [string]$service.DisplayName
                State       = [string]$service.State
                StartName   = [string]$service.StartName
                FileName    = ''
                Path        = ''
                Entry       = ''
                MatchSource = if ($isNameOnlyCandidate) { 'NameOnly' } else { 'Account' }
                IsOptInCandidate = $isNameOnlyCandidate
                MatchReason = if ($isNameOnlyCandidate) { 'Service name contains the instance token; Log On As does not.' } else { 'Log On As contains the instance token.' }
            }
            Write-Output (
                '@@SERVICE_TARGET@@' +
                ($discoveryRecord | ConvertTo-Json -Compress)
            )
        }
        foreach ($listTarget in $gcsListTargets) {
            $discoveryRecord = [PSCustomObject]@{
                TargetType  = 'GCSListEntry'
                Server      = $server
                Name        = [string]$listTarget.Entry
                DisplayName = [string]$listTarget.FileName
                State       = 'Remove line'
                StartName   = ''
                FileName    = [string]$listTarget.FileName
                Path        = [string]$listTarget.Path
                Entry       = [string]$listTarget.Entry
                MatchSource = 'GCSList'
                IsOptInCandidate = $false
                MatchReason = 'GCS list entry contains the instance token.'
            }
            Write-Output (
                '@@SERVICE_TARGET@@' +
                ($discoveryRecord | ConvertTo-Json -Compress)
            )
        }
    } else {
        if ($matches.Count -gt 0) {
            Write-Host ("[{0}] Matching services:" -f $server) -ForegroundColor Yellow
            foreach ($service in $matches) {
                Write-Host ("  {0} ({1}) | State={2} | As={3}" -f $service.Name, $service.DisplayName, $service.State, $service.StartName)
            }
        }
        if ($gcsListTargets.Count -gt 0) {
            Write-Host ("[{0}] Matching GCS service-list entries:" -f $server) -ForegroundColor Yellow
            foreach ($listTarget in $gcsListTargets) {
                Write-Host ("  {0} | Entry={1}" -f $listTarget.Path, $listTarget.Entry)
            }
        }
    }

    if ($DiscoveryOnly) {
        continue
    }

    $serverServiceDeletionFailed = $false
    foreach ($service in $matches) {
        $target = '{0}\{1}' -f $server, $service.Name
        $action = "Stop and delete service '$($service.DisplayName)' running as '$($service.StartName)'"
        if (-not $PSCmdlet.ShouldProcess($target, $action)) {
            continue
        }
        if (-not (Confirm-DestructiveAction -Message ("Delete service '{0}' from {1}?" -f $service.Name, $server) -Force:$Force)) {
            Write-Host ("Skipped service '{0}'." -f $service.Name)
            $serverServiceDeletionFailed = $true
            continue
        }

        $stopResult = Stop-RemoteServiceSafe -Server $server -ServiceName $service.Name -TimeoutSec $StopTimeoutSec
        if (-not $stopResult.Success) {
            Write-Warning ("[{0}] Service '{1}' was not deleted because stopping failed: {2}" -f $server, $service.Name, $stopResult.Message)
            $hadFailures = $true
            $serverServiceDeletionFailed = $true
            continue
        }

        $removeResult = Remove-RemoteService -Server $server -ServiceName $service.Name
        if ($removeResult.Success) {
            Write-Host ("[{0}] {1}: {2}" -f $server, $service.Name, $removeResult.Message) -ForegroundColor Green
        } else {
            Write-Warning ("[{0}] Failed to remove service '{1}': {2}" -f $server, $service.Name, $removeResult.Message)
            $hadFailures = $true
            $serverServiceDeletionFailed = $true
        }
    }

    if ($serverServiceDeletionFailed -and -not $WhatIfPreference) {
        if ($gcsListTargets.Count -gt 0) {
            Write-Warning ("[{0}] GCS service-list cleanup was skipped because one or more confirmed services were not deleted." -f $server)
        }
        continue
    }

    foreach ($fileGroup in @($gcsListTargets | Group-Object Path)) {
        $path = [string]$fileGroup.Name
        $entries = @($fileGroup.Group | ForEach-Object { [string]$_.Entry } | Select-Object -Unique)
        $action = "Remove $($entries.Count) confirmed instance service entr$(if ($entries.Count -eq 1) { 'y' } else { 'ies' }) from GCS list"
        if (-not $PSCmdlet.ShouldProcess($path, $action)) {
            continue
        }
        if (-not (Confirm-DestructiveAction -Message ("Remove {0} confirmed entr{1} from {2}?" -f $entries.Count, $(if ($entries.Count -eq 1) { 'y' } else { 'ies' }), $path) -Force:$Force)) {
            Write-Host ("Skipped GCS service-list cleanup: {0}" -f $path)
            continue
        }

        $cleanupResult = Remove-ConfirmedGCSServiceListEntries `
            -Path $path `
            -ConfirmedEntries $entries `
            -Token $LogOnAsToken
        if ($cleanupResult.Success) {
            Write-Host ("[{0}] {1}: {2}" -f $server, $path, $cleanupResult.Message) -ForegroundColor Green
            foreach ($removedEntry in @($cleanupResult.Removed)) {
                Write-Host ("  Removed list entry: {0}" -f $removedEntry) -ForegroundColor Green
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$cleanupResult.BackupPath)) {
                Write-Host ("  Backup: {0}" -f $cleanupResult.BackupPath)
            }
        } else {
            Write-Warning ("[{0}] Failed to update {1}: {2}" -f $server, $path, $cleanupResult.Message)
            $hadFailures = $true
        }
    }
}

if ($hadFailures) {
    if ($DiscoveryOnly) {
        Write-Warning 'Service discovery completed with one or more failures.'
    } else {
        Write-Warning 'Service decommission completed with one or more failures.'
    }
    exit 2
}

if ($DiscoveryOnly) {
    Write-Host 'Service discovery completed successfully.'
} elseif ($WhatIfPreference) {
    Write-Host 'Service preview completed successfully.'
} else {
    Write-Host 'Service decommission completed successfully.'
}
