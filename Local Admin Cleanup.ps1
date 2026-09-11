#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string[]]$ComputerName,

    [string]$Instance,

    [switch]$DiscoveryOnly,

    [switch]$EmitDiscoveryJson,

    [string]$TargetManifestBase64,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$script:AdministratorsSid = 'S-1-5-32-544'
$script:AuditRecords = [System.Collections.Generic.List[object]]::new()
$runTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$auditDirectory = Join-Path -Path $PSScriptRoot -ChildPath 'Logs'
$script:AuditPath = Join-Path -Path $auditDirectory -ChildPath "LocalAdmins_Audit_$runTimestamp.csv"

function ConvertTo-ServerList {
    param([string[]]$InputNames)

    @(
        $InputNames |
            ForEach-Object { $_ -split '[,;\s]+' } |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ } |
            Sort-Object -Unique
    )
}

function Get-BuiltinAdministratorsGroup {
    param([Parameter(Mandatory)][string]$Server)

    # Resolve the well-known SID on the target so localized group names work.
    $computer = [ADSI]("WinNT://{0},computer" -f $Server)
    foreach ($child in @($computer.psbase.Children)) {
        try {
            if ([string]$child.SchemaClassName -ine 'Group') { continue }

            $sidBytes = [byte[]]$child.psbase.InvokeGet('objectSID')
            if (-not $sidBytes) { continue }

            $sid = [System.Security.Principal.SecurityIdentifier]::new($sidBytes, 0)
            if ($sid.Value -eq $script:AdministratorsSid) {
                $groupName = [string]$child.Name
                return [PSCustomObject]@{
                    Name = $groupName
                    Path = "WinNT://$Server/$groupName,group"
                }
            }
        } catch {
            continue
        }
    }

    throw "The built-in Administrators group ($script:AdministratorsSid) was not found."
}

function Get-GroupMembers {
    param([Parameter(Mandatory)][string]$GroupPath)

    $group = [ADSI]$GroupPath
    foreach ($member in @($group.psbase.Invoke('Members'))) {
        $memberPath = [string]$member.GetType().InvokeMember(
            'ADsPath', 'GetProperty', $null, $member, $null
        )
        $memberClass = [string]$member.GetType().InvokeMember(
            'Class', 'GetProperty', $null, $member, $null
        )
        $memberName = [string]$member.GetType().InvokeMember(
            'Name', 'GetProperty', $null, $member, $null
        )

        [PSCustomObject]@{
            Name  = $memberName
            Class = $memberClass
            Path  = $memberPath
        }
    }
}

function Test-GroupMemberPath {
    param(
        [Parameter(Mandatory)][string]$GroupPath,
        [Parameter(Mandatory)][string]$MemberPath
    )

    foreach ($member in @(Get-GroupMembers -GroupPath $GroupPath)) {
        if ([System.StringComparer]::OrdinalIgnoreCase.Equals($member.Path, $MemberPath)) {
            return $true
        }
    }
    return $false
}

function ConvertFrom-LocalAdminTargetManifest {
    param(
        [Parameter(Mandatory)][string]$Base64,
        [Parameter(Mandatory)][string[]]$AllowedServers
    )

    $json = [System.Text.Encoding]::UTF8.GetString(
        [System.Convert]::FromBase64String($Base64)
    )
    $decodedManifest = $json | ConvertFrom-Json -ErrorAction Stop

    # Windows PowerShell 5.1 returns a JSON array as one non-enumerated
    # Object[] pipeline item. Pipe the stored value again so each confirmed
    # target remains a separate record instead of joining every Server value.
    $decoded = @(
        $decodedManifest |
            ForEach-Object { $_ }
    )
    if ($decoded.Count -eq 0) {
        throw 'The local-admin target manifest is empty.'
    }
    $records = [System.Collections.Generic.List[object]]::new()

    foreach ($record in $decoded) {
        $server = [string]$record.Server
        $name = [string]$record.Name
        $path = [string]$record.Path

        if (
            [string]::IsNullOrWhiteSpace($server) -or
            [string]::IsNullOrWhiteSpace($name) -or
            [string]::IsNullOrWhiteSpace($path)
        ) {
            throw 'Every local-admin target must contain Server, Name, and Path.'
        }

        $serverAllowed = $false
        foreach ($allowedServer in $AllowedServers) {
            if ([System.StringComparer]::OrdinalIgnoreCase.Equals($allowedServer, $server)) {
                $serverAllowed = $true
                break
            }
        }
        if (-not $serverAllowed) {
            throw "Target manifest server '$server' is not in ComputerName."
        }

        $duplicate = $false
        foreach ($existing in $records) {
            if (
                [System.StringComparer]::OrdinalIgnoreCase.Equals($existing.Server, $server) -and
                [System.StringComparer]::OrdinalIgnoreCase.Equals($existing.Path, $path)
            ) {
                $duplicate = $true
                break
            }
        }
        if (-not $duplicate) {
            $records.Add([PSCustomObject]@{
                Server = $server
                Name   = $name
                Path   = $path
            })
        }
    }

    return $records
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

function Add-AuditRecord {
    param(
        [string]$Server,
        [string]$Group,
        [string]$Name,
        [string]$Class,
        [string]$Path,
        [string]$Action,
        [string]$Result,
        [string]$ErrorMessage
    )

    $script:AuditRecords.Add([PSCustomObject]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Server    = $Server
        Group     = $Group
        Name      = $Name
        Class     = $Class
        Path      = $Path
        Action    = $Action
        Result    = $Result
        Error     = $ErrorMessage
    })
}

function Save-AuditReport {
    try {
        if (-not (Test-Path -LiteralPath $auditDirectory)) {
            $null = New-Item -Path $auditDirectory -ItemType Directory -Force
        }
        $script:AuditRecords |
            Export-Csv -LiteralPath $script:AuditPath -NoTypeInformation -Encoding UTF8
        Write-Host ("Audit report: {0}" -f $script:AuditPath)
    } catch {
        Write-Warning ("Unable to save audit report '{0}': {1}" -f $script:AuditPath, $_.Exception.Message)
    }
}

if (-not $ComputerName -or $ComputerName.Count -eq 0) {
    $ComputerName = @((Read-Host 'Enter servers (comma, space, or semicolon separated)'))
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
    $Instance = Read-Host 'Enter the instance token to match in group names'
}
if ([string]::IsNullOrWhiteSpace($Instance)) {
    Write-Error 'No instance token provided.'
    exit 1
}
$Instance = $Instance.Trim()

if ($EmitDiscoveryJson -and -not $DiscoveryOnly) {
    Write-Error '-EmitDiscoveryJson can be used only with -DiscoveryOnly.'
    exit 1
}

$manifest = @()
$hasManifest = -not [string]::IsNullOrWhiteSpace($TargetManifestBase64)
if ($hasManifest) {
    try {
        $manifest = @(
            ConvertFrom-LocalAdminTargetManifest `
                -Base64 $TargetManifestBase64 `
                -AllowedServers $servers
        )
    } catch {
        Write-Error ("Invalid local-admin target manifest: {0}" -f $_.Exception.Message)
        exit 1
    }
}

Write-Host ("Servers: {0}" -f ($servers -join ', '))
Write-Host ("Instance token: {0} (literal, case-insensitive substring)" -f $Instance)
if ($WhatIfPreference) {
    Write-Host '[WhatIf] Preview mode is ON. Local Administrators memberships will not be changed.'
}

$hadFailures = $false
$totalMatches = 0
$totalRemoved = 0

foreach ($server in $servers) {
    Write-Host ("Discovering local Administrators group members on {0}..." -f $server)
    try {
        $groupInfo = Get-BuiltinAdministratorsGroup -Server $server
        $matchingMembers = @(
            Get-GroupMembers -GroupPath $groupInfo.Path |
                Where-Object {
                    $_.Class -ieq 'Group' -and
                    $_.Name.IndexOf($Instance, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
                } |
                Sort-Object Path -Unique
        )
    } catch {
        $hadFailures = $true
        $message = $_.Exception.Message
        Write-Warning ("ERROR: Discovery failed on {0}: {1}" -f $server, $message)
        Add-AuditRecord -Server $server -Group $null -Name $null -Class $null `
            -Path $null -Action 'Discover' -Result 'Failed' -ErrorMessage $message
        continue
    }

    $totalMatches += $matchingMembers.Count
    Write-Host ("Found {0} matching group principal(s) on {1}." -f $matchingMembers.Count, $server)

    foreach ($member in $matchingMembers) {
        Add-AuditRecord -Server $server -Group $groupInfo.Name -Name $member.Name `
            -Class $member.Class -Path $member.Path -Action 'Discover' `
            -Result 'Matched' -ErrorMessage $null

        if ($DiscoveryOnly -and $EmitDiscoveryJson) {
            $record = [PSCustomObject]@{
                Server = $server
                Group  = $groupInfo.Name
                Name   = $member.Name
                Class  = $member.Class
                Path   = $member.Path
            }
            Write-Output ('@@LOCAL_ADMIN_TARGET@@' + ($record | ConvertTo-Json -Compress))
        } elseif ($DiscoveryOnly) {
            Write-Host ("MATCH: {0}\{1} -> {2}" -f $server, $groupInfo.Name, $member.Path)
        }
    }

    if ($DiscoveryOnly) { continue }

    $targets = @($matchingMembers)
    if ($hasManifest) {
        $serverManifest = @(
            $manifest |
                Where-Object {
                    [System.StringComparer]::OrdinalIgnoreCase.Equals($_.Server, $server)
                }
        )
        $targets = @(
            foreach ($requestedTarget in $serverManifest) {
                $liveTarget = @(
                    $matchingMembers |
                        Where-Object {
                            [System.StringComparer]::OrdinalIgnoreCase.Equals($_.Path, $requestedTarget.Path)
                        }
                ) | Select-Object -First 1

                if ($null -eq $liveTarget) {
                    $hadFailures = $true
                    $message = 'Confirmed target is no longer a matching group member.'
                    Write-Warning ("Skipped {0} on {1}: {2}" -f $requestedTarget.Name, $server, $message)
                    Add-AuditRecord -Server $server -Group $groupInfo.Name `
                        -Name $requestedTarget.Name -Class 'Group' -Path $requestedTarget.Path `
                        -Action 'Remove' -Result 'Skipped' -ErrorMessage $message
                    continue
                }
                $liveTarget
            }
        )
    }

    foreach ($targetMember in $targets) {
        $displayTarget = '{0}\{1}\{2}' -f $server, $groupInfo.Name, $targetMember.Name
        $action = "Remove group principal '$($targetMember.Name)' from local Administrators"

        if (-not $PSCmdlet.ShouldProcess($displayTarget, $action)) {
            $result = if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' }
            Add-AuditRecord -Server $server -Group $groupInfo.Name `
                -Name $targetMember.Name -Class $targetMember.Class -Path $targetMember.Path `
                -Action 'Remove' -Result $result -ErrorMessage $null
            continue
        }

        if (-not (Confirm-DestructiveAction -Message ("Remove {0}?" -f $displayTarget) -Force:$Force)) {
            Write-Host ("Skipped {0}." -f $displayTarget)
            Add-AuditRecord -Server $server -Group $groupInfo.Name `
                -Name $targetMember.Name -Class $targetMember.Class -Path $targetMember.Path `
                -Action 'Remove' -Result 'Cancelled' -ErrorMessage $null
            continue
        }

        try {
            if (-not (Test-GroupMemberPath -GroupPath $groupInfo.Path -MemberPath $targetMember.Path)) {
                throw 'The principal was no longer a member when removal began.'
            }

            $group = [ADSI]$groupInfo.Path
            $group.Remove($targetMember.Path)

            if (Test-GroupMemberPath -GroupPath $groupInfo.Path -MemberPath $targetMember.Path) {
                throw 'The remove call completed, but the principal is still a member.'
            }

            $totalRemoved++
            Write-Host ("Removed and verified: {0}" -f $displayTarget) -ForegroundColor Green
            Add-AuditRecord -Server $server -Group $groupInfo.Name `
                -Name $targetMember.Name -Class $targetMember.Class -Path $targetMember.Path `
                -Action 'Remove' -Result 'Success' -ErrorMessage $null
        } catch {
            $hadFailures = $true
            $message = $_.Exception.Message
            Write-Warning ("ERROR: Failed to remove {0}: {1}" -f $displayTarget, $message)
            Add-AuditRecord -Server $server -Group $groupInfo.Name `
                -Name $targetMember.Name -Class $targetMember.Class -Path $targetMember.Path `
                -Action 'Remove' -Result 'Failed' -ErrorMessage $message
        }
    }
}

# The combined SCAN already records discovery in the session activity log and
# exposes the matches in its read-only result grid.  Create this separate audit
# CSV only for an actual cleanup or a WhatIf cleanup preview.
if (-not $DiscoveryOnly) {
    if ($script:AuditRecords.Count -eq 0) {
        Add-AuditRecord -Server $null -Group $null -Name $Instance -Class 'Group' `
            -Path $null -Action 'Discover' -Result 'NoMatches' -ErrorMessage $null
    }
    Save-AuditReport
}

if ($hadFailures) {
    if ($DiscoveryOnly) {
        Write-Warning 'Local-admin discovery completed with one or more failures. No cleanup was attempted.'
    } else {
        Write-Warning 'Local-admin cleanup completed with one or more failures or changed targets.'
    }
    exit 2
}

if ($DiscoveryOnly) {
    Write-Host ("Local-admin discovery completed successfully. {0} target(s) found." -f $totalMatches)
} elseif ($WhatIfPreference) {
    Write-Host ("Local-admin preview completed successfully. {0} target(s) found." -f $totalMatches)
} else {
    Write-Host ("Local-admin cleanup completed successfully. {0} group principal(s) removed." -f $totalRemoved)
}
