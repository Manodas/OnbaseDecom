#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string[]]$ComputerName,

    [string[]]$FolderName,

    [string[]]$EncodedFolderName,

    [switch]$Force
)

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

function ConvertTo-FolderList {
    param([string[]]$InputNames)

    @(
        $InputNames |
            ForEach-Object { $_ -split '[,;\r\n]+' } |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ } |
            Select-Object -Unique
    )
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

function Clear-RestrictiveAttributes {
    param([Parameter(Mandatory)][string]$Path)

    $clearMask = [int](
        [System.IO.FileAttributes]::ReadOnly -bor
        [System.IO.FileAttributes]::Hidden -bor
        [System.IO.FileAttributes]::System
    )

    $items = @(
        Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
        Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
    )
    foreach ($item in $items) {
        try {
            $newAttributes = [int]$item.Attributes -band (-bnot $clearMask)
            $item.Attributes = [System.IO.FileAttributes]$newAttributes
        } catch {}
    }
}

$remoteDeleteScript = {
    param([string]$LeafName)

    function Clear-RestrictiveAttributes {
        param([string]$Path)

        $clearMask = [int](
            [System.IO.FileAttributes]::ReadOnly -bor
            [System.IO.FileAttributes]::Hidden -bor
            [System.IO.FileAttributes]::System
        )
        $items = @(
            Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
            Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
        )
        foreach ($item in $items) {
            try {
                $newAttributes = [int]$item.Attributes -band (-bnot $clearMask)
                $item.Attributes = [System.IO.FileAttributes]$newAttributes
            } catch {}
        }
    }

    $basePath = 'M:\OBOL\Clients'
    $target = Join-Path -Path $basePath -ChildPath $LeafName

    try {
        if (-not (Test-Path -LiteralPath $target -ErrorAction Stop)) {
            return [PSCustomObject]@{
                Success = $true
                Message = 'Folder was already absent.'
            }
        }

        Clear-RestrictiveAttributes -Path $target
        $lastError = ''
        $aclAdjusted = $false

        for ($attempt = 1; $attempt -le 4; $attempt++) {
            try {
                Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction Stop
                if (Test-Path -LiteralPath $target -ErrorAction Stop) {
                    throw 'The folder still exists after Remove-Item completed.'
                }

                return [PSCustomObject]@{
                    Success = $true
                    Message = 'Folder removed and absence verified.'
                }
            } catch {
                $lastError = $_.Exception.Message

                if (-not $aclAdjusted) {
                    try {
                        & takeown.exe /f $target /r /d y 2>&1 | Out-Null
                        & icacls.exe $target /grant '*S-1-5-32-544:F' /t /c 2>&1 | Out-Null
                    } catch {}
                    $aclAdjusted = $true
                }

                if ($attempt -lt 4) {
                    Start-Sleep -Seconds ([int][Math]::Pow(2, $attempt))
                }
            }
        }

        return [PSCustomObject]@{
            Success = $false
            Message = $lastError
        }
    } catch {
        return [PSCustomObject]@{
            Success = $false
            Message = $_.Exception.Message
        }
    }
}

function Remove-ClientFolder {
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$LeafName
    )

    try {
        $remoteResult = Invoke-Command -ComputerName $Server -ScriptBlock $remoteDeleteScript -ArgumentList $LeafName -ErrorAction Stop
        if ($remoteResult.Success) {
            return [PSCustomObject]@{
                Success = $true
                Method  = 'Remote'
                Message = $remoteResult.Message
            }
        }

        Write-Warning ("[{0}] Remote removal failed for '{1}': {2}. Trying UNC." -f $Server, $LeafName, $remoteResult.Message)
    } catch {
        Write-Warning ("[{0}] Remote execution failed for '{1}': {2}. Trying UNC." -f $Server, $LeafName, $_.Exception.Message)
    }

    $uncPath = "\\$Server\m$\OBOL\Clients\$LeafName"
    try {
        if (-not (Test-Path -LiteralPath $uncPath -ErrorAction Stop)) {
            return [PSCustomObject]@{
                Success = $true
                Method  = 'UNC'
                Message = 'Folder was already absent.'
            }
        }

        Clear-RestrictiveAttributes -Path $uncPath
        Remove-Item -LiteralPath $uncPath -Recurse -Force -ErrorAction Stop
        if (Test-Path -LiteralPath $uncPath -ErrorAction Stop) {
            throw 'The folder still exists after Remove-Item completed.'
        }

        return [PSCustomObject]@{
            Success = $true
            Method  = 'UNC'
            Message = 'Folder removed and absence verified.'
        }
    } catch {
        return [PSCustomObject]@{
            Success = $false
            Method  = 'UNC'
            Message = $_.Exception.Message
        }
    }
}

if (-not $ComputerName -or $ComputerName.Count -eq 0) {
    $ComputerName = @((Read-Host 'Enter service and application servers'))
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

if ($EncodedFolderName -and $EncodedFolderName.Count -gt 0) {
    try {
        $folders = @(
            $EncodedFolderName |
                ForEach-Object { $_ -split '[,;\s]+' } |
                Where-Object { $_ } |
                ForEach-Object {
                    [System.Text.Encoding]::UTF8.GetString(
                        [System.Convert]::FromBase64String($_)
                    )
                } |
                ForEach-Object { $_.Trim() } |
                Where-Object { $_ } |
                Select-Object -Unique
        )
    } catch {
        Write-Error ("An encoded folder selection is invalid: {0}" -f $_.Exception.Message)
        exit 1
    }
} else {
    if (-not $FolderName -or $FolderName.Count -eq 0) {
        $FolderName = @((Read-Host 'Enter exact folder names (comma or semicolon separated)'))
    }
    $folders = @(ConvertTo-FolderList -InputNames $FolderName)
}

if ($folders.Count -eq 0) {
    Write-Error 'No folder names provided.'
    exit 1
}

foreach ($folder in $folders) {
    if (
        $folder -in @('.', '..') -or
        $folder -match '[\\/:*?"<>|]' -or
        $folder.EndsWith('.') -or
        $folder.EndsWith(' ')
    ) {
        Write-Error ("Invalid folder leaf name '{0}'. Enter exact names beneath M:\OBOL\Clients only." -f $folder)
        exit 1
    }
}

Write-Host ("Folder servers: {0}" -f ($servers -join ', '))
Write-Host ("Exact folder names: {0}" -f ($folders -join ', '))
if ($WhatIfPreference) {
    Write-Host '[WhatIf] Preview mode is ON. No folder will be deleted.'
}

$hadFailures = $false

foreach ($server in $servers) {
    Write-Host ("==== {0} :: Client Folders ====" -f $server) -ForegroundColor Cyan

    foreach ($folder in $folders) {
        $target = "\\$server\m$\OBOL\Clients\$folder"
        $action = "Recursively delete client folder '$folder'"
        if (-not $PSCmdlet.ShouldProcess($target, $action)) {
            continue
        }
        if (-not (Confirm-DestructiveAction -Message ("Delete '{0}' from {1}?" -f $folder, $server) -Force:$Force)) {
            Write-Host ("Skipped folder '{0}' on {1}." -f $folder, $server)
            continue
        }

        $result = Remove-ClientFolder -Server $server -LeafName $folder
        if ($result.Success) {
            Write-Host ("[{0}] {1} ({2}): {3}" -f $server, $folder, $result.Method, $result.Message) -ForegroundColor Green
        } else {
            Write-Warning ("[{0}] Failed to remove '{1}' via {2}: {3}" -f $server, $folder, $result.Method, $result.Message)
            $hadFailures = $true
        }
    }
}

if ($hadFailures) {
    Write-Warning 'Folder deletion completed with one or more failures.'
    exit 2
}

if ($WhatIfPreference) {
    Write-Host 'Folder deletion preview completed successfully.'
} else {
    Write-Host 'Folder deletion completed successfully.'
}
