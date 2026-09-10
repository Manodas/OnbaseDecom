#Requires -Version 5.1

<#
Enumerate and optionally remove IIS sites and application pools whose names
match an instance token across multiple servers. Matching child IIS
applications under preserved shared sites are also removed.

Safety behavior:
- -WhatIf performs discovery only and makes no IIS changes.
- Every destructive operation requires a Y response that defaults to N.
- -Force bypasses the custom Y/N prompt, but -WhatIf still takes precedence.
- Default Web Site is never selected for removal.
- Matching non-root child applications under preserved non-default sites are
  selected for removal without deleting their parent sites or physical files.
- Application pools are selected independently when they match the instance
  token.
- Application pools still referenced by any IIS application are not removed
  unless -AllowSharedAppPoolRemoval is explicitly supplied.
- Deletion is reported as successful only after the IIS object is verified
  absent.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string[]]$ComputerName,

    [string]$Instance,

    [ValidateSet('Token', 'Exact', 'Contains')]
    [string]$MatchMode = 'Token',

    [switch]$AllowSharedAppPoolRemoval,

    [switch]$DiscoveryOnly,

    [switch]$EmitDiscoveryJson,

    [string]$TargetManifestBase64,

    [switch]$Force
)

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

function Get-IISApplicationKey {
    param(
        [string]$SiteName,
        [string]$ApplicationPath
    )

    '{0}|{1}' -f $SiteName.Trim(), $ApplicationPath.Trim()
}

function Test-IISConsumerInRemovalPlan {
    param(
        [string]$SiteName,
        [string]$ApplicationPath,
        [string[]]$MatchedSiteNames,
        [string[]]$MatchedApplicationKeys
    )

    if ($MatchedSiteNames -contains $SiteName) {
        return $true
    }

    $consumerKey = Get-IISApplicationKey `
        -SiteName $SiteName `
        -ApplicationPath $ApplicationPath
    return ($MatchedApplicationKeys -contains $consumerKey)
}

function ConvertFrom-IISTargetManifest {
    param(
        [Parameter(Mandatory)][string]$Base64,
        [Parameter(Mandatory)][string[]]$AllowedServers
    )

    $manifestJson = [System.Text.Encoding]::UTF8.GetString(
        [System.Convert]::FromBase64String($Base64)
    )
    $decodedManifest = $manifestJson | ConvertFrom-Json -ErrorAction Stop

    # Windows PowerShell 5.1 can surface a JSON array as one Object[] object.
    # Enumerate explicitly so server and target properties remain scalar.
    $records = @(
        $decodedManifest |
            ForEach-Object { $_ }
    )
    if ($records.Count -eq 0) {
        throw 'The IIS target manifest is empty.'
    }

    foreach ($record in $records) {
        $server = [string]$record.Server
        $targetType = [string]$record.TargetType
        if ([string]::IsNullOrWhiteSpace($server)) {
            throw 'Every IIS target manifest record must contain Server.'
        }
        if ($AllowedServers -notcontains $server) {
            throw "IIS target manifest server '$server' is not in ComputerName."
        }
        if ($targetType -notin @('Site', 'Application', 'AppPool')) {
            throw "Invalid IIS target type '$targetType'."
        }

        switch ($targetType) {
            'Site' {
                if ([string]::IsNullOrWhiteSpace([string]$record.Name)) {
                    throw 'Every Site target must contain Name.'
                }
            }
            'Application' {
                if (
                    [string]::IsNullOrWhiteSpace([string]$record.SiteName) -or
                    [string]::IsNullOrWhiteSpace([string]$record.ApplicationPath) -or
                    [string]$record.ApplicationPath -eq '/'
                ) {
                    throw 'Every Application target must contain a non-root SiteName and ApplicationPath.'
                }
            }
            'AppPool' {
                if ([string]::IsNullOrWhiteSpace([string]$record.Name)) {
                    throw 'Every AppPool target must contain Name.'
                }
            }
        }
    }

    return $records
}

$discoveryScript = {
    param(
        [string]$InstanceToken,
        [string]$NameMatchMode
    )

    function Initialize-IISProvider {
        $importError = $null
        try {
            Import-Module WebAdministration -ErrorAction Stop
        } catch {
            $importError = $_.Exception.Message
        }

        $iisDrive = Get-PSDrive -Name IIS -ErrorAction SilentlyContinue
        if ($null -eq $iisDrive) {
            if ($importError) {
                throw "WebAdministration could not be loaded and the IIS: provider is unavailable. $importError"
            }
            throw 'The IIS: provider is unavailable.'
        }
    }

    function New-IISServerManager {
        if ($null -eq ('Microsoft.Web.Administration.ServerManager' -as [type])) {
            $assemblyCandidates = @(
                (Join-Path $env:windir 'System32\inetsrv\Microsoft.Web.Administration.dll')
                (Join-Path $env:windir 'Sysnative\inetsrv\Microsoft.Web.Administration.dll')
            )
            $assemblyPath = @(
                $assemblyCandidates |
                    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf }
            ) | Select-Object -First 1

            if ([string]::IsNullOrWhiteSpace([string]$assemblyPath)) {
                throw 'Microsoft.Web.Administration.dll was not found. Confirm that the IIS Management Scripts and Tools feature is installed.'
            }

            Add-Type -Path $assemblyPath -ErrorAction Stop
        }

        return New-Object Microsoft.Web.Administration.ServerManager -ErrorAction Stop
    }

    function Test-IISNameMatch {
        param(
            [string]$Name,
            [string]$Token,
            [string]$Mode
        )

        if ([string]::IsNullOrWhiteSpace($Name)) { return $false }

        switch ($Mode) {
            'Exact' {
                return $Name.Equals($Token, [System.StringComparison]::OrdinalIgnoreCase)
            }
            'Contains' {
                return $Name.IndexOf($Token, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            }
            default {
                $pattern = '(^|[^A-Za-z0-9]){0}($|[^A-Za-z0-9])' -f [regex]::Escape($Token)
                return [regex]::IsMatch($Name, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            }
        }
    }

    $serverManager = $null
    try {
        Initialize-IISProvider
        $serverManager = New-IISServerManager

        $allSites = @(Get-ChildItem -LiteralPath 'IIS:\Sites' -ErrorAction Stop)
        $allPools = @(Get-ChildItem -LiteralPath 'IIS:\AppPools' -ErrorAction Stop)
        $managedSites = @($serverManager.Sites)
        $siteRecords = @()
        $applicationRecords = @()
        $poolRecords = @()
        $consumerRecords = @()
        $matchedManagedSiteNames = @(
            $managedSites |
                Where-Object {
                    $managedSiteName = [string]$_.Name
                    -not $managedSiteName.Equals(
                        'Default Web Site',
                        [System.StringComparison]::OrdinalIgnoreCase
                    ) -and
                    (Test-IISNameMatch -Name $managedSiteName -Token $InstanceToken -Mode $NameMatchMode)
                } |
                ForEach-Object { [string]$_.Name }
        )

        # Read every site's pool assignments for dependency protection.
        # Child applications are targets only under preserved non-default sites.
        foreach ($managedSite in $managedSites) {
            $managedSiteName = [string]$managedSite.Name
            $applications = @($managedSite.Applications)
            if ($applications.Count -eq 0) {
                throw "Unable to enumerate applications for IIS site '$managedSiteName'."
            }

            foreach ($application in $applications) {
                $applicationPath = [string]$application.Path
                $poolName = [string]$application.ApplicationPoolName
                if ([string]::IsNullOrWhiteSpace($applicationPath)) {
                    throw "An application in site '$managedSiteName' has no readable path."
                }
                if ([string]::IsNullOrWhiteSpace($poolName)) {
                    throw "Application '$applicationPath' in site '$managedSiteName' has no readable application-pool assignment."
                }

                $consumerRecords += [PSCustomObject]@{
                    SiteName       = $managedSiteName
                    ApplicationPath = $applicationPath
                    AppPoolName    = $poolName
                }

                $isDefaultSite = $managedSiteName.Equals(
                    'Default Web Site',
                    [System.StringComparison]::OrdinalIgnoreCase
                )
                $parentSiteIsTargeted = $matchedManagedSiteNames -contains $managedSiteName
                $applicationMatches = Test-IISNameMatch `
                    -Name $applicationPath `
                    -Token $InstanceToken `
                    -Mode $NameMatchMode

                if (
                    -not $isDefaultSite -and
                    -not $parentSiteIsTargeted -and
                    $applicationPath -ne '/' -and
                    $applicationMatches
                ) {
                    $physicalPath = '[Unable to determine]'
                    try {
                        $rootDirectory = @(
                            $application.VirtualDirectories |
                                Where-Object { $_.Path -eq '/' }
                        )[0]
                        if ($null -ne $rootDirectory) {
                            $physicalPath = [string]$rootDirectory.PhysicalPath
                        }
                    } catch {}

                    $applicationRecords += [PSCustomObject]@{
                        SiteName        = $managedSiteName
                        ApplicationPath = $applicationPath
                        AppPoolName     = $poolName
                        PhysicalPath    = $physicalPath
                    }
                }
            }
        }

        foreach ($site in $allSites) {
            $siteName = [string]$site.Name
            if ($siteName.Equals('Default Web Site', [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            if (Test-IISNameMatch -Name $siteName -Token $InstanceToken -Mode $NameMatchMode) {
                $bindingText = ''
                try {
                    $bindingText = @(
                        $site.Bindings.Collection |
                            ForEach-Object {
                                '{0}://{1}' -f $_.protocol, $_.bindingInformation
                            }
                    ) -join '; '
                } catch {
                    $bindingText = '[Unable to format bindings]'
                }

                $physicalPath = [string]$site.PhysicalPath
                if ([string]::IsNullOrWhiteSpace($physicalPath)) {
                    try {
                        $managedSite = @(
                            $managedSites |
                                Where-Object {
                                    ([string]$_.Name).Equals(
                                        [string]$site.Name,
                                        [System.StringComparison]::OrdinalIgnoreCase
                                    )
                                }
                        ) | Select-Object -First 1
                        if ($null -eq $managedSite) {
                            throw "Microsoft.Web.Administration did not return site '$($site.Name)'."
                        }

                        $rootApplication = @($managedSite.Applications | Where-Object { $_.Path -eq '/' })[0]
                        $rootDirectory = @($rootApplication.VirtualDirectories | Where-Object { $_.Path -eq '/' })[0]
                        $physicalPath = [string]$rootDirectory.PhysicalPath
                    } catch {
                        $physicalPath = '[Unable to determine]'
                    }
                }

                $siteRecords += [PSCustomObject]@{
                    Name         = $siteName
                    State        = [string]$site.State
                    PhysicalPath = $physicalPath
                    Bindings     = $bindingText
                }
            }
        }

        foreach ($pool in $allPools) {
            $poolName = [string]$pool.Name
            $matchesInstance = Test-IISNameMatch -Name $poolName -Token $InstanceToken -Mode $NameMatchMode

            if ($matchesInstance) {
                $poolIdentity = '[Unable to determine]'
                try {
                    $managedPool = @(
                        $serverManager.ApplicationPools |
                            Where-Object {
                                ([string]$_.Name).Equals(
                                    $poolName,
                                    [System.StringComparison]::OrdinalIgnoreCase
                                )
                            }
                    ) | Select-Object -First 1
                    if ($null -eq $managedPool) {
                        throw "Microsoft.Web.Administration did not return application pool '$poolName'."
                    }

                    $identityType = [string]$managedPool.ProcessModel.IdentityType
                    $identityUser = [string]$managedPool.ProcessModel.UserName
                    $poolIdentity = if ([string]::IsNullOrWhiteSpace($identityUser)) {
                        $identityType
                    } else {
                        $identityUser
                    }
                } catch {
                    $poolIdentity = '[Unable to determine]'
                }
                $identityMatchesToken = (
                    -not [string]::IsNullOrWhiteSpace($poolIdentity) -and
                    $poolIdentity.IndexOf($InstanceToken, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
                )
                $poolRecords += [PSCustomObject]@{
                    Name                  = $poolName
                    State                 = [string]$pool.State
                    ManagedRuntimeVersion = [string]$pool.managedRuntimeVersion
                    PipelineMode          = [string]$pool.managedPipelineMode
                    PoolIdentity          = $poolIdentity
                    IdentityMatchesToken  = $identityMatchesToken
                    IsOptInCandidate      = -not $identityMatchesToken
                }
            }
        }

        [PSCustomObject]@{
            Success   = $true
            Error     = ''
            Sites     = @($siteRecords)
            Applications = @($applicationRecords)
            AppPools  = @($poolRecords)
            Consumers = @($consumerRecords)
        }
    } catch {
        [PSCustomObject]@{
            Success   = $false
            Error     = $_.Exception.Message
            Sites     = @()
            Applications = @()
            AppPools  = @()
            Consumers = @()
        }
    } finally {
        if ($null -ne $serverManager) {
            $serverManager.Dispose()
        }
    }
}

$removeSiteScript = {
    param([string]$SiteName)

    function Initialize-IISProvider {
        try { Import-Module WebAdministration -ErrorAction Stop } catch {}
        if ($null -eq (Get-PSDrive -Name IIS -ErrorAction SilentlyContinue)) {
            throw 'The IIS: provider is unavailable.'
        }
    }

    try {
        Initialize-IISProvider
        $sitePath = Join-Path -Path 'IIS:\Sites' -ChildPath $SiteName

        if (-not (Test-Path -LiteralPath $sitePath -ErrorAction Stop)) {
            return [PSCustomObject]@{
                Success = $true
                Name    = $SiteName
                Message = 'Site was already absent.'
            }
        }

        $siteItem = Get-Item -LiteralPath $sitePath -ErrorAction Stop
        if ([string]$siteItem.State -ne 'Stopped') {
            if (Get-Command Stop-Website -ErrorAction SilentlyContinue) {
                Stop-Website -Name $SiteName -ErrorAction Stop | Out-Null
            } else {
                $siteItem.Stop() | Out-Null
            }

            $siteItem = Get-Item -LiteralPath $sitePath -ErrorAction Stop
            if ([string]$siteItem.State -ne 'Stopped') {
                throw "The site did not reach the Stopped state. Current state: $($siteItem.State)"
            }
        }

        if (Get-Command Remove-Website -ErrorAction SilentlyContinue) {
            Remove-Website -Name $SiteName -Confirm:$false -ErrorAction Stop
        } else {
            Remove-Item -LiteralPath $sitePath -Recurse -Force -ErrorAction Stop
        }

        if (Test-Path -LiteralPath $sitePath -ErrorAction Stop) {
            throw 'The site still exists after the removal command completed.'
        }

        return [PSCustomObject]@{
            Success = $true
            Name    = $SiteName
            Message = 'Site removed and absence verified.'
        }
    } catch {
        return [PSCustomObject]@{
            Success = $false
            Name    = $SiteName
            Message = $_.Exception.Message
        }
    }
}

$removeApplicationScript = {
    param(
        [string]$SiteName,
        [string]$ApplicationPath
    )

    function New-IISServerManager {
        if ($null -eq ('Microsoft.Web.Administration.ServerManager' -as [type])) {
            $assemblyCandidates = @(
                (Join-Path $env:windir 'System32\inetsrv\Microsoft.Web.Administration.dll')
                (Join-Path $env:windir 'Sysnative\inetsrv\Microsoft.Web.Administration.dll')
            )
            $assemblyPath = @(
                $assemblyCandidates |
                    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf }
            ) | Select-Object -First 1

            if ([string]::IsNullOrWhiteSpace([string]$assemblyPath)) {
                throw 'Microsoft.Web.Administration.dll was not found. Confirm that the IIS Management Scripts and Tools feature is installed.'
            }

            Add-Type -Path $assemblyPath -ErrorAction Stop
        }

        return New-Object Microsoft.Web.Administration.ServerManager -ErrorAction Stop
    }

    if (
        [string]::IsNullOrWhiteSpace($SiteName) -or
        [string]::IsNullOrWhiteSpace($ApplicationPath) -or
        $ApplicationPath -eq '/'
    ) {
        return [PSCustomObject]@{
            Success = $false
            SiteName = $SiteName
            ApplicationPath = $ApplicationPath
            Message = 'Refusing to remove an invalid or root IIS application target.'
        }
    }

    $serverManager = $null
    try {
        $serverManager = New-IISServerManager
        $site = @(
            $serverManager.Sites |
                Where-Object {
                    ([string]$_.Name).Equals(
                        $SiteName,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                }
        ) | Select-Object -First 1

        if ($null -eq $site) {
            return [PSCustomObject]@{
                Success = $true
                SiteName = $SiteName
                ApplicationPath = $ApplicationPath
                Message = 'Parent site was already absent; application is absent.'
            }
        }

        $application = @(
            $site.Applications |
                Where-Object {
                    ([string]$_.Path).Equals(
                        $ApplicationPath,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                }
        ) | Select-Object -First 1

        if ($null -eq $application) {
            return [PSCustomObject]@{
                Success = $true
                SiteName = $SiteName
                ApplicationPath = $ApplicationPath
                Message = 'Application was already absent.'
            }
        }

        [void]$site.Applications.Remove($application)
        $serverManager.CommitChanges()
    } catch {
        return [PSCustomObject]@{
            Success = $false
            SiteName = $SiteName
            ApplicationPath = $ApplicationPath
            Message = $_.Exception.Message
        }
    } finally {
        if ($null -ne $serverManager) {
            $serverManager.Dispose()
        }
    }

    $verificationManager = $null
    try {
        $verificationManager = New-IISServerManager
        $verificationSite = @(
            $verificationManager.Sites |
                Where-Object {
                    ([string]$_.Name).Equals(
                        $SiteName,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                }
        ) | Select-Object -First 1

        if ($null -ne $verificationSite) {
            $remainingApplication = @(
                $verificationSite.Applications |
                    Where-Object {
                        ([string]$_.Path).Equals(
                            $ApplicationPath,
                            [System.StringComparison]::OrdinalIgnoreCase
                        )
                    }
            ) | Select-Object -First 1

            if ($null -ne $remainingApplication) {
                throw 'The IIS application still exists after CommitChanges completed.'
            }
        }

        return [PSCustomObject]@{
            Success = $true
            SiteName = $SiteName
            ApplicationPath = $ApplicationPath
            Message = 'Application removed and absence verified; parent site and physical files were preserved.'
        }
    } catch {
        return [PSCustomObject]@{
            Success = $false
            SiteName = $SiteName
            ApplicationPath = $ApplicationPath
            Message = $_.Exception.Message
        }
    } finally {
        if ($null -ne $verificationManager) {
            $verificationManager.Dispose()
        }
    }
}

$removePoolScript = {
    param(
        [string]$PoolName,
        [bool]$AllowInUse
    )

    function Initialize-IISProvider {
        try { Import-Module WebAdministration -ErrorAction Stop } catch {}
        if ($null -eq (Get-PSDrive -Name IIS -ErrorAction SilentlyContinue)) {
            throw 'The IIS: provider is unavailable.'
        }
    }

    function New-IISServerManager {
        if ($null -eq ('Microsoft.Web.Administration.ServerManager' -as [type])) {
            $assemblyCandidates = @(
                (Join-Path $env:windir 'System32\inetsrv\Microsoft.Web.Administration.dll')
                (Join-Path $env:windir 'Sysnative\inetsrv\Microsoft.Web.Administration.dll')
            )
            $assemblyPath = @(
                $assemblyCandidates |
                    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf }
            ) | Select-Object -First 1

            if ([string]::IsNullOrWhiteSpace([string]$assemblyPath)) {
                throw 'Microsoft.Web.Administration.dll was not found. Confirm that the IIS Management Scripts and Tools feature is installed.'
            }

            Add-Type -Path $assemblyPath -ErrorAction Stop
        }

        return New-Object Microsoft.Web.Administration.ServerManager -ErrorAction Stop
    }

    try {
        Initialize-IISProvider
        $liveConsumers = @()
        $serverManager = $null

        try {
            $serverManager = New-IISServerManager
            foreach ($site in @($serverManager.Sites)) {
                $applications = @($site.Applications)
                if ($applications.Count -eq 0) {
                    throw "Unable to enumerate applications for IIS site '$($site.Name)'."
                }

                foreach ($application in $applications) {
                    if ([string]$application.ApplicationPoolName -eq $PoolName) {
                        $liveConsumers += '{0}{1}' -f $site.Name, $application.Path
                    }
                }
            }
        } finally {
            if ($null -ne $serverManager) {
                $serverManager.Dispose()
            }
        }

        if ($liveConsumers.Count -gt 0 -and -not $AllowInUse) {
            return [PSCustomObject]@{
                Success   = $false
                Blocked   = $true
                Name      = $PoolName
                Consumers = @($liveConsumers)
                Message   = 'Application pool is still referenced by one or more IIS applications.'
            }
        }

        $poolPath = Join-Path -Path 'IIS:\AppPools' -ChildPath $PoolName
        if (-not (Test-Path -LiteralPath $poolPath -ErrorAction Stop)) {
            return [PSCustomObject]@{
                Success   = $true
                Blocked   = $false
                Name      = $PoolName
                Consumers = @($liveConsumers)
                Message   = 'Application pool was already absent.'
            }
        }

        $poolItem = Get-Item -LiteralPath $poolPath -ErrorAction Stop
        if ([string]$poolItem.State -ne 'Stopped') {
            if (Get-Command Stop-WebAppPool -ErrorAction SilentlyContinue) {
                Stop-WebAppPool -Name $PoolName -ErrorAction Stop | Out-Null
            } else {
                $poolItem.Stop() | Out-Null
            }

            $poolItem = Get-Item -LiteralPath $poolPath -ErrorAction Stop
            if ([string]$poolItem.State -ne 'Stopped') {
                throw "The application pool did not reach the Stopped state. Current state: $($poolItem.State)"
            }
        }

        if (Get-Command Remove-WebAppPool -ErrorAction SilentlyContinue) {
            Remove-WebAppPool -Name $PoolName -Confirm:$false -ErrorAction Stop
        } else {
            Remove-Item -LiteralPath $poolPath -Recurse -Force -ErrorAction Stop
        }

        if (Test-Path -LiteralPath $poolPath -ErrorAction Stop) {
            throw 'The application pool still exists after the removal command completed.'
        }

        return [PSCustomObject]@{
            Success   = $true
            Blocked   = $false
            Name      = $PoolName
            Consumers = @($liveConsumers)
            Message   = 'Application pool removed and absence verified.'
        }
    } catch {
        return [PSCustomObject]@{
            Success   = $false
            Blocked   = $false
            Name      = $PoolName
            Consumers = @()
            Message   = $_.Exception.Message
        }
    }
}

if (-not $ComputerName -or $ComputerName.Count -eq 0) {
    $serverInput = Read-Host 'Enter server names (comma, space, or semicolon separated)'
    $ComputerName = @($serverInput)
}

$servers = @(ConvertTo-ServerList -InputNames $ComputerName)
if ($servers.Count -eq 0) {
    Write-Error 'No servers provided.'
    exit 1
}

foreach ($server in $servers) {
    if ($server -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        Write-Error ("Invalid server name '{0}'. Enter a host name, FQDN, or IPv4 address only." -f $server)
        exit 1
    }
}

if ([string]::IsNullOrWhiteSpace($Instance)) {
    $Instance = Read-Host 'Enter the instance name to search for'
}
if ([string]::IsNullOrWhiteSpace($Instance)) {
    Write-Error 'No instance name provided.'
    exit 1
}
$Instance = $Instance.Trim()

if ($EmitDiscoveryJson -and -not $DiscoveryOnly) {
    Write-Error '-EmitDiscoveryJson can be used only with -DiscoveryOnly.'
    exit 1
}

$targetManifest = @()
$hasTargetManifest = -not [string]::IsNullOrWhiteSpace($TargetManifestBase64)
if ($hasTargetManifest) {
    try {
        $targetManifest = @(
            ConvertFrom-IISTargetManifest `
                -Base64 $TargetManifestBase64 `
                -AllowedServers $servers
        )
    } catch {
        Write-Error ("Invalid IIS target manifest: {0}" -f $_.Exception.Message)
        exit 1
    }
}

Write-Host ("Servers: {0}" -f ($servers -join ', '))
Write-Host ("Instance: {0} | Match mode: {1}" -f $Instance, $MatchMode)
if ($WhatIfPreference) {
    Write-Host '[WhatIf] Preview mode is ON. IIS configuration will not be changed.'
}
if ($AllowSharedAppPoolRemoval) {
    Write-Warning 'Shared application-pool removal override is enabled.'
}

$hadFailures = $false

foreach ($server in $servers) {
    Write-Host ("==== {0} ====" -f $server) -ForegroundColor Cyan

    try {
        $result = Invoke-Command -ComputerName $server -ScriptBlock $discoveryScript -ArgumentList $Instance, $MatchMode -ErrorAction Stop
    } catch {
        Write-Warning ("Failed to query {0}: {1}" -f $server, $_.Exception.Message)
        $hadFailures = $true
        continue
    }

    if (-not $result.Success) {
        Write-Warning ("IIS discovery failed on {0}: {1}. No deletion was attempted." -f $server, $result.Error)
        $hadFailures = $true
        continue
    }

    $sites = @($result.Sites | Where-Object { $null -ne $_ })
    $applications = @($result.Applications | Where-Object { $null -ne $_ })
    $pools = @($result.AppPools | Where-Object { $null -ne $_ })
    $consumers = @($result.Consumers | Where-Object { $null -ne $_ })

    if ($hasTargetManifest) {
        $serverTargets = @(
            $targetManifest |
                Where-Object { ([string]$_.Server) -eq $server }
        )
        $confirmedSiteNames = @(
            $serverTargets |
                Where-Object { [string]$_.TargetType -eq 'Site' } |
                ForEach-Object { [string]$_.Name } |
                Select-Object -Unique
        )
        $confirmedApplicationKeys = @(
            $serverTargets |
                Where-Object { [string]$_.TargetType -eq 'Application' } |
                ForEach-Object {
                    Get-IISApplicationKey `
                        -SiteName ([string]$_.SiteName) `
                        -ApplicationPath ([string]$_.ApplicationPath)
                } |
                Select-Object -Unique
        )
        $confirmedPoolNames = @(
            $serverTargets |
                Where-Object { [string]$_.TargetType -eq 'AppPool' } |
                ForEach-Object { [string]$_.Name } |
                Select-Object -Unique
        )

        $sites = @(
            $sites |
                Where-Object { $confirmedSiteNames -contains ([string]$_.Name) }
        )
        $applications = @(
            $applications |
                Where-Object {
                    $applicationKey = Get-IISApplicationKey `
                        -SiteName ([string]$_.SiteName) `
                        -ApplicationPath ([string]$_.ApplicationPath)
                    $confirmedApplicationKeys -contains $applicationKey
                }
        )
        $pools = @(
            $pools |
                Where-Object { $confirmedPoolNames -contains ([string]$_.Name) }
        )
    } elseif (-not $DiscoveryOnly) {
        # Identity-mismatch pools are opt-in only. Destructive and WhatIf runs
        # require an exact manifest generated from checked scan candidates.
        $pools = @($pools | Where-Object { -not [bool]$_.IsOptInCandidate })
    }

    if ($sites.Count -eq 0 -and $applications.Count -eq 0 -and $pools.Count -eq 0) {
        Write-Host ("No IIS sites, child applications, or application pools matched '{0}' on {1}." -f $Instance, $server) -ForegroundColor DarkYellow
        continue
    }

    if ($sites.Count -gt 0) {
        Write-Host ("Matched sites on {0}:" -f $server) -ForegroundColor Green
        foreach ($site in $sites) {
            Write-Host ("  {0} | State={1} | Path={2} | Bindings={3}" -f $site.Name, $site.State, $site.PhysicalPath, $site.Bindings)
        }
    } else {
        Write-Host ("No matching sites on {0}." -f $server) -ForegroundColor DarkYellow
    }

    if ($applications.Count -gt 0) {
        Write-Host ("Matched child applications under preserved sites on {0}:" -f $server) -ForegroundColor Green
        foreach ($application in $applications) {
            Write-Host (
                "  Site={0} | Application={1} | Pool={2} | Path={3}" -f
                    $application.SiteName,
                    $application.ApplicationPath,
                    $application.AppPoolName,
                    $application.PhysicalPath
            )
        }
    } else {
        Write-Host ("No matching child applications under preserved sites on {0}." -f $server) -ForegroundColor DarkYellow
    }

    if ($pools.Count -gt 0) {
        Write-Host ("Matched application pools on {0}:" -f $server) -ForegroundColor Green
        foreach ($pool in $pools) {
            $poolConsumers = @($consumers | Where-Object { $_.AppPoolName -eq $pool.Name })
            $consumerText = if ($poolConsumers.Count -eq 0) {
                '[unused]'
            } else {
                @($poolConsumers | ForEach-Object { '{0}{1}' -f $_.SiteName, $_.ApplicationPath }) -join ', '
            }
            Write-Host ("  {0} | State={1} | Identity={2} | Runtime={3} | Pipeline={4} | Consumers={5}" -f $pool.Name, $pool.State, $pool.PoolIdentity, $pool.ManagedRuntimeVersion, $pool.PipelineMode, $consumerText)
        }
    } else {
        Write-Host ("No matching application pools on {0}." -f $server) -ForegroundColor DarkYellow
    }

    $matchedSiteNames = @($sites | ForEach-Object { $_.Name })
    $matchedApplicationKeys = @(
        $applications |
            ForEach-Object {
                Get-IISApplicationKey `
                    -SiteName $_.SiteName `
                    -ApplicationPath $_.ApplicationPath
            }
    )

    if ($DiscoveryOnly) {
        if ($EmitDiscoveryJson) {
            foreach ($site in $sites) {
                $record = [PSCustomObject]@{
                    Server          = $server
                    TargetType      = 'Site'
                    Name            = [string]$site.Name
                    SiteName        = [string]$site.Name
                    ApplicationPath = ''
                    AppPoolName     = ''
                    State           = [string]$site.State
                    Planned         = $true
                    Details         = 'Path={0}; Bindings={1}' -f $site.PhysicalPath, $site.Bindings
                    BlockedReason   = ''
                }
                Write-Output ('@@IIS_TARGET@@' + ($record | ConvertTo-Json -Compress))
            }

            foreach ($application in $applications) {
                $record = [PSCustomObject]@{
                    Server          = $server
                    TargetType      = 'Application'
                    Name            = '{0}{1}' -f $application.SiteName, $application.ApplicationPath
                    SiteName        = [string]$application.SiteName
                    ApplicationPath = [string]$application.ApplicationPath
                    AppPoolName     = [string]$application.AppPoolName
                    State           = ''
                    Planned         = $true
                    Details         = 'Physical path={0}' -f $application.PhysicalPath
                    BlockedReason   = ''
                }
                Write-Output ('@@IIS_TARGET@@' + ($record | ConvertTo-Json -Compress))
            }

            foreach ($pool in $pools) {
                $poolConsumers = @($consumers | Where-Object { $_.AppPoolName -eq $pool.Name })
                $externalConsumers = @(
                    $poolConsumers |
                        Where-Object {
                            -not (Test-IISConsumerInRemovalPlan `
                                -SiteName $_.SiteName `
                                -ApplicationPath $_.ApplicationPath `
                                -MatchedSiteNames $matchedSiteNames `
                                -MatchedApplicationKeys $matchedApplicationKeys)
                        }
                )
                $isBlocked = $externalConsumers.Count -gt 0 -and -not $AllowSharedAppPoolRemoval
                $consumerText = if ($poolConsumers.Count -eq 0) {
                    '[unused]'
                } else {
                    @($poolConsumers | ForEach-Object { '{0}{1}' -f $_.SiteName, $_.ApplicationPath }) -join ', '
                }
                $blockedReason = if ($isBlocked) {
                    'Used outside removal plan: {0}' -f (
                        @($externalConsumers | ForEach-Object { '{0}{1}' -f $_.SiteName, $_.ApplicationPath }) -join ', '
                    )
                } else {
                    ''
                }

                $record = [PSCustomObject]@{
                    Server          = $server
                    TargetType      = 'AppPool'
                    Name            = [string]$pool.Name
                    SiteName        = ''
                    ApplicationPath = ''
                    AppPoolName     = [string]$pool.Name
                    State           = [string]$pool.State
                    Planned         = -not $isBlocked
                    Details         = 'Identity={0}; Runtime={1}; Pipeline={2}; Consumers={3}' -f $pool.PoolIdentity, $pool.ManagedRuntimeVersion, $pool.PipelineMode, $consumerText
                    BlockedReason   = $blockedReason
                    PoolIdentity    = [string]$pool.PoolIdentity
                    IdentityMatchesToken = [bool]$pool.IdentityMatchesToken
                    IsOptInCandidate = [bool]$pool.IsOptInCandidate
                    MatchSource     = if ([bool]$pool.IsOptInCandidate) { 'NameOnly' } else { 'Identity' }
                    MatchReason     = if ([bool]$pool.IsOptInCandidate) { 'Application-pool name contains the instance token; pool identity does not.' } else { 'Application-pool identity contains the instance token.' }
                }
                Write-Output ('@@IIS_TARGET@@' + ($record | ConvertTo-Json -Compress))
            }
        }

        continue
    }

    foreach ($site in $sites) {
        $target = '{0}\IIS:\Sites\{1}' -f $server, $site.Name
        $action = "Stop and remove IIS site '$($site.Name)'"
        if (-not $PSCmdlet.ShouldProcess($target, $action)) {
            continue
        }
        if (-not (Confirm-DestructiveAction -Message ("Remove site '{0}' from {1}?" -f $site.Name, $server) -Force:$Force)) {
            Write-Host ("Skipped site '{0}'." -f $site.Name)
            continue
        }

        try {
            $removeResult = Invoke-Command -ComputerName $server -ScriptBlock $removeSiteScript -ArgumentList $site.Name -ErrorAction Stop
        } catch {
            Write-Warning ("Failed to remove site '{0}' on {1}: {2}" -f $site.Name, $server, $_.Exception.Message)
            $hadFailures = $true
            continue
        }

        if ($removeResult.Success) {
            Write-Host ("[{0}] {1}: {2}" -f $server, $site.Name, $removeResult.Message) -ForegroundColor Green
        } else {
            Write-Warning ("Failed to remove site '{0}' on {1}: {2}" -f $site.Name, $server, $removeResult.Message)
            $hadFailures = $true
        }
    }

    $applicationsToRemove = @(
        $applications |
            Sort-Object `
                @{ Expression = { ([string]$_.ApplicationPath).Length }; Descending = $true },
                SiteName
    )

    foreach ($application in $applicationsToRemove) {
        $applicationDisplayName = '{0}{1}' -f $application.SiteName, $application.ApplicationPath
        $target = '{0}\IIS:\Sites\{1}\Applications{2}' -f $server, $application.SiteName, $application.ApplicationPath
        $action = "Remove IIS child application '$applicationDisplayName' while preserving its parent site and physical files"
        if (-not $PSCmdlet.ShouldProcess($target, $action)) {
            continue
        }
        if (-not (Confirm-DestructiveAction -Message ("Remove IIS application '{0}' from {1}?" -f $applicationDisplayName, $server) -Force:$Force)) {
            Write-Host ("Skipped IIS application '{0}'." -f $applicationDisplayName)
            continue
        }

        try {
            $removeResult = Invoke-Command `
                -ComputerName $server `
                -ScriptBlock $removeApplicationScript `
                -ArgumentList $application.SiteName, $application.ApplicationPath `
                -ErrorAction Stop
        } catch {
            Write-Warning ("Failed to remove IIS application '{0}' on {1}: {2}" -f $applicationDisplayName, $server, $_.Exception.Message)
            $hadFailures = $true
            continue
        }

        if ($removeResult.Success) {
            Write-Host ("[{0}] {1}: {2}" -f $server, $applicationDisplayName, $removeResult.Message) -ForegroundColor Green
        } else {
            Write-Warning ("Failed to remove IIS application '{0}' on {1}: {2}" -f $applicationDisplayName, $server, $removeResult.Message)
            $hadFailures = $true
        }
    }

    foreach ($pool in $pools) {
        $poolConsumers = @($consumers | Where-Object { $_.AppPoolName -eq $pool.Name })
        $externalConsumers = @(
            $poolConsumers |
                Where-Object {
                    -not (Test-IISConsumerInRemovalPlan `
                        -SiteName $_.SiteName `
                        -ApplicationPath $_.ApplicationPath `
                        -MatchedSiteNames $matchedSiteNames `
                        -MatchedApplicationKeys $matchedApplicationKeys)
                }
        )

        if ($externalConsumers.Count -gt 0 -and -not $AllowSharedAppPoolRemoval) {
            $consumerText = @($externalConsumers | ForEach-Object { '{0}{1}' -f $_.SiteName, $_.ApplicationPath }) -join ', '
            Write-Warning ("Blocked removal of app pool '{0}' on {1}; it is used by applications outside the removal plan: {2}" -f $pool.Name, $server, $consumerText)
            $hadFailures = $true
            continue
        }

        $target = '{0}\IIS:\AppPools\{1}' -f $server, $pool.Name
        $action = if ($poolConsumers.Count -gt 0) {
            "Stop and remove IIS application pool '$($pool.Name)' after rechecking $($poolConsumers.Count) consumer(s)"
        } else {
            "Stop and remove unused IIS application pool '$($pool.Name)'"
        }

        if (-not $PSCmdlet.ShouldProcess($target, $action)) {
            continue
        }
        if (-not (Confirm-DestructiveAction -Message ("Remove application pool '{0}' from {1}?" -f $pool.Name, $server) -Force:$Force)) {
            Write-Host ("Skipped application pool '{0}'." -f $pool.Name)
            continue
        }

        try {
            $removeResult = Invoke-Command -ComputerName $server -ScriptBlock $removePoolScript -ArgumentList $pool.Name, ([bool]$AllowSharedAppPoolRemoval) -ErrorAction Stop
        } catch {
            Write-Warning ("Failed to remove application pool '{0}' on {1}: {2}" -f $pool.Name, $server, $_.Exception.Message)
            $hadFailures = $true
            continue
        }

        if ($removeResult.Success) {
            Write-Host ("[{0}] {1}: {2}" -f $server, $pool.Name, $removeResult.Message) -ForegroundColor Green
        } elseif ($removeResult.Blocked) {
            Write-Warning ("Blocked removal of app pool '{0}' on {1}; live consumers: {2}" -f $pool.Name, $server, (@($removeResult.Consumers) -join ', '))
            $hadFailures = $true
        } else {
            Write-Warning ("Failed to remove application pool '{0}' on {1}: {2}" -f $pool.Name, $server, $removeResult.Message)
            $hadFailures = $true
        }
    }
}

if ($hadFailures) {
    if ($DiscoveryOnly) {
        Write-Warning 'IIS discovery completed with one or more failures. No deletion was attempted.'
    } else {
        Write-Warning 'Completed with one or more discovery, dependency, or removal failures. Review the warnings above.'
    }
    exit 2
}

if ($DiscoveryOnly) {
    Write-Host 'IIS discovery completed successfully.'
} elseif ($WhatIfPreference) {
    Write-Host 'Preview completed successfully.'
} else {
    Write-Host 'Completed successfully.'
}
