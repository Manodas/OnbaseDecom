#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$RenderPreviewPath,

    [string]$RenderFolderPickerPreviewPath,

    [string]$RenderServiceConfirmationPreviewPath,

    [string]$RenderIISConfirmationPreviewPath,

    [string]$RenderLocalAdminConfirmationPreviewPath,

    [string]$RenderScanResultsPreviewPath,

    [string]$RenderScanResultsPreviewTab,

    [switch]$RenderScanResultsWidePreview,

    [string]$RenderCandidateSelectionPreviewPath,

    [string]$RenderODBCConfirmationPreviewPath
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$serviceScript = Join-Path -Path $PSScriptRoot -ChildPath 'Service deletion.ps1'
$iisScript = Join-Path -Path $PSScriptRoot -ChildPath 'IIS site and pool deletion.ps1'
$folderScript = Join-Path -Path $PSScriptRoot -ChildPath 'Folder deletion.ps1'
$localAdminScript = Join-Path -Path $PSScriptRoot -ChildPath 'Local Admin Cleanup.ps1'
$scanScript = Join-Path -Path $PSScriptRoot -ChildPath 'Scan inventory.ps1'
$odbcScript = Join-Path -Path $PSScriptRoot -ChildPath 'ODBC System DSN Cleanup.ps1'

foreach ($requiredFile in @($serviceScript, $iisScript, $folderScript, $localAdminScript, $scanScript, $odbcScript)) {
    if (-not (Test-Path -LiteralPath $requiredFile)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Required application file is missing:`r`n$requiredFile",
            'Decommissioning Tool',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'POST DECOM TOOL'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1100, 980)
$form.MinimumSize = New-Object System.Drawing.Size(940, 780)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = 'POST DECOM TOOL'
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 18)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 55)
$titleLabel.Location = New-Object System.Drawing.Point(20, 14)
$titleLabel.AutoSize = $true
$form.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = 'Run service, IIS, local-admin, and client-folder decommissioning from one controlled interface.'
$subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(90, 100, 115)
$subtitleLabel.Location = New-Object System.Drawing.Point(23, 50)
$subtitleLabel.AutoSize = $true
$form.Controls.Add($subtitleLabel)

$inputGroup = New-Object System.Windows.Forms.GroupBox
$inputGroup.Text = 'Inputs'
$inputGroup.Location = New-Object System.Drawing.Point(20, 78)
$inputGroup.Size = New-Object System.Drawing.Size(1040, 258)
$inputGroup.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Left -bor
    [System.Windows.Forms.AnchorStyles]::Right
)
$form.Controls.Add($inputGroup)

$serviceServerLabel = New-Object System.Windows.Forms.Label
$serviceServerLabel.Text = 'Service servers'
$serviceServerLabel.Location = New-Object System.Drawing.Point(18, 28)
$serviceServerLabel.AutoSize = $true
$inputGroup.Controls.Add($serviceServerLabel)

$serviceServersTextBox = New-Object System.Windows.Forms.TextBox
$serviceServersTextBox.Location = New-Object System.Drawing.Point(18, 50)
$serviceServersTextBox.Size = New-Object System.Drawing.Size(490, 62)
$serviceServersTextBox.Multiline = $true
$serviceServersTextBox.ScrollBars = 'Vertical'
$serviceServersTextBox.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Left
)
$inputGroup.Controls.Add($serviceServersTextBox)

$applicationServerLabel = New-Object System.Windows.Forms.Label
$applicationServerLabel.Text = 'Application / IIS servers'
$applicationServerLabel.Location = New-Object System.Drawing.Point(528, 28)
$applicationServerLabel.AutoSize = $true
$inputGroup.Controls.Add($applicationServerLabel)

$applicationServersTextBox = New-Object System.Windows.Forms.TextBox
$applicationServersTextBox.Location = New-Object System.Drawing.Point(528, 50)
$applicationServersTextBox.Size = New-Object System.Drawing.Size(490, 62)
$applicationServersTextBox.Multiline = $true
$applicationServersTextBox.ScrollBars = 'Vertical'
$applicationServersTextBox.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Left -bor
    [System.Windows.Forms.AnchorStyles]::Right
)
$inputGroup.Controls.Add($applicationServersTextBox)

$serviceAccountLabel = New-Object System.Windows.Forms.Label
$serviceAccountLabel.Text = "Service 'Log On As' match"
$serviceAccountLabel.Location = New-Object System.Drawing.Point(398, 124)
$serviceAccountLabel.AutoSize = $true
$inputGroup.Controls.Add($serviceAccountLabel)

$serviceAccountTextBox = New-Object System.Windows.Forms.TextBox
$serviceAccountTextBox.Location = New-Object System.Drawing.Point(398, 146)
$serviceAccountTextBox.Size = New-Object System.Drawing.Size(330, 25)
$serviceAccountTextBox.ReadOnly = $true
$serviceAccountTextBox.TabStop = $false
$serviceAccountTextBox.Text = 'Account name contains instance token'
$serviceAccountTextBox.BackColor = [System.Drawing.Color]::FromArgb(235, 238, 242)
$serviceAccountTextBox.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Left
)
$inputGroup.Controls.Add($serviceAccountTextBox)

$instanceLabel = New-Object System.Windows.Forms.Label
$instanceLabel.Text = 'Instance token (IIS, services, and local admins)'
$instanceLabel.Location = New-Object System.Drawing.Point(18, 124)
$instanceLabel.AutoSize = $true
$inputGroup.Controls.Add($instanceLabel)

$instanceTextBox = New-Object System.Windows.Forms.TextBox
$instanceTextBox.Location = New-Object System.Drawing.Point(18, 146)
$instanceTextBox.Size = New-Object System.Drawing.Size(360, 25)
$inputGroup.Controls.Add($instanceTextBox)

$matchModeLabel = New-Object System.Windows.Forms.Label
$matchModeLabel.Text = 'IIS match mode'
$matchModeLabel.Location = New-Object System.Drawing.Point(748, 124)
$matchModeLabel.AutoSize = $true
$inputGroup.Controls.Add($matchModeLabel)

$matchModeComboBox = New-Object System.Windows.Forms.ComboBox
$matchModeComboBox.Location = New-Object System.Drawing.Point(748, 146)
$matchModeComboBox.Size = New-Object System.Drawing.Size(270, 25)
$matchModeComboBox.DropDownStyle = 'DropDownList'
[void]$matchModeComboBox.Items.AddRange(@('Token', 'Exact', 'Contains'))
$matchModeComboBox.SelectedIndex = 0
$inputGroup.Controls.Add($matchModeComboBox)

$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Text = 'Client folders under M:\OBOL\Clients on the service and application servers'
$folderLabel.Location = New-Object System.Drawing.Point(18, 184)
$folderLabel.AutoSize = $true
$inputGroup.Controls.Add($folderLabel)

$selectFoldersButton = New-Object System.Windows.Forms.Button
$selectFoldersButton.Text = 'Choose Client Folders...'
$selectFoldersButton.Location = New-Object System.Drawing.Point(18, 207)
$selectFoldersButton.Size = New-Object System.Drawing.Size(210, 32)
$selectFoldersButton.FlatStyle = 'Flat'
$selectFoldersButton.BackColor = [System.Drawing.Color]::FromArgb(55, 65, 81)
$selectFoldersButton.ForeColor = [System.Drawing.Color]::White
$selectFoldersButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$inputGroup.Controls.Add($selectFoldersButton)

$folderSelectionSummary = New-Object System.Windows.Forms.Label
$folderSelectionSummary.Text = 'No folders selected'
$folderSelectionSummary.Location = New-Object System.Drawing.Point(240, 207)
$folderSelectionSummary.Size = New-Object System.Drawing.Size(778, 32)
$folderSelectionSummary.BorderStyle = 'FixedSingle'
$folderSelectionSummary.TextAlign = 'MiddleLeft'
$folderSelectionSummary.AutoEllipsis = $true
$folderSelectionSummary.Padding = New-Object System.Windows.Forms.Padding(8, 0, 8, 0)
$folderSelectionSummary.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Left -bor
    [System.Windows.Forms.AnchorStyles]::Right
)
$inputGroup.Controls.Add($folderSelectionSummary)

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Location = New-Object System.Drawing.Point(20, 346)
$buttonPanel.Size = New-Object System.Drawing.Size(1040, 48)
$buttonPanel.FlowDirection = 'LeftToRight'
$buttonPanel.WrapContents = $false
$buttonPanel.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Left -bor
    [System.Windows.Forms.AnchorStyles]::Right
)
$form.Controls.Add($buttonPanel)

function New-ActionButton {
    param(
        [string]$Text,
        [System.Drawing.Color]$BackColor,
        [int]$Width
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Size = New-Object System.Drawing.Size($Width, 38)
    $button.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
    $button.FlatStyle = 'Flat'
    $button.FlatAppearance.BorderSize = 0
    $button.BackColor = $BackColor
    $button.ForeColor = [System.Drawing.Color]::White
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $button
}

$deleteServicesButton = New-ActionButton -Text 'Delete Services' -BackColor ([System.Drawing.Color]::FromArgb(185, 28, 28)) -Width 105
$iisButton = New-ActionButton -Text 'IIS Pool and Site Decom' -BackColor ([System.Drawing.Color]::FromArgb(180, 83, 9)) -Width 155
$folderButton = New-ActionButton -Text 'Folder Deletion' -BackColor ([System.Drawing.Color]::FromArgb(190, 24, 93)) -Width 110
$localAdminButton = New-ActionButton -Text 'Local Admin Cleanup' -BackColor ([System.Drawing.Color]::FromArgb(126, 34, 206)) -Width 140
$odbcButton = New-ActionButton -Text 'ODBC DSN Cleanup' -BackColor ([System.Drawing.Color]::FromArgb(3, 105, 161)) -Width 135
$scanButton = New-ActionButton -Text 'SCAN' -BackColor ([System.Drawing.Color]::FromArgb(8, 145, 178)) -Width 70
$whatIfButton = New-ActionButton -Text 'What If (All)' -BackColor ([System.Drawing.Color]::FromArgb(37, 99, 235)) -Width 95
$cancelButton = New-ActionButton -Text 'Cancel Action' -BackColor ([System.Drawing.Color]::FromArgb(75, 85, 99)) -Width 100
$cancelButton.Enabled = $false

[void]$buttonPanel.Controls.Add($deleteServicesButton)
[void]$buttonPanel.Controls.Add($iisButton)
[void]$buttonPanel.Controls.Add($folderButton)
[void]$buttonPanel.Controls.Add($localAdminButton)
[void]$buttonPanel.Controls.Add($odbcButton)
[void]$buttonPanel.Controls.Add($scanButton)
[void]$buttonPanel.Controls.Add($whatIfButton)
[void]$buttonPanel.Controls.Add($cancelButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = 'Ready'
$statusLabel.Location = New-Object System.Drawing.Point(23, 402)
$statusLabel.Size = New-Object System.Drawing.Size(700, 22)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
$statusLabel.AutoEllipsis = $true
$form.Controls.Add($statusLabel)

$statusToolTip = New-Object System.Windows.Forms.ToolTip
$statusToolTip.AutoPopDelay = 30000
$statusToolTip.InitialDelay = 400
$statusToolTip.ReshowDelay = 100
$statusToolTip.ShowAlways = $true
$statusToolTip.SetToolTip($statusLabel, $statusLabel.Text)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(760, 400)
$progressBar.Size = New-Object System.Drawing.Size(300, 18)
$progressBar.Style = 'Continuous'
$progressBar.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Right
)
$form.Controls.Add($progressBar)

$resultsSplitContainer = New-Object System.Windows.Forms.SplitContainer
$resultsSplitContainer.Location = New-Object System.Drawing.Point(20, 430)
$resultsSplitContainer.Size = New-Object System.Drawing.Size(1040, 505)
$resultsSplitContainer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$resultsSplitContainer.SplitterDistance = 350
$resultsSplitContainer.Panel1MinSize = 180
$resultsSplitContainer.Panel2MinSize = 145
$resultsSplitContainer.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$resultsSplitContainer.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Left -bor
    [System.Windows.Forms.AnchorStyles]::Right
)
$form.Controls.Add($resultsSplitContainer)

$scanGroup = New-Object System.Windows.Forms.GroupBox
$scanGroup.Text = 'Scan results (read-only)'
$scanGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
$resultsSplitContainer.Panel1.Controls.Add($scanGroup)

$scanSummary = New-Object System.Windows.Forms.Label
$scanSummary.Text = 'Click SCAN to display matching resources. Scanning does not change configuration.'
$scanSummary.Location = New-Object System.Drawing.Point(12, 20)
$scanSummary.Size = New-Object System.Drawing.Size(810, 22)
$scanSummary.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
$scanSummary.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Left -bor
    [System.Windows.Forms.AnchorStyles]::Right
)
$scanGroup.Controls.Add($scanSummary)

$exportScanButton = New-Object System.Windows.Forms.Button
$exportScanButton.Text = 'Export Scan Results to Excel...'
$exportScanButton.Location = New-Object System.Drawing.Point(823, 16)
$exportScanButton.Size = New-Object System.Drawing.Size(205, 27)
$exportScanButton.Enabled = $false
$exportScanButton.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Right
)
$scanGroup.Controls.Add($exportScanButton)

$scanTabs = New-Object System.Windows.Forms.TabControl
$scanTabs.Location = New-Object System.Drawing.Point(12, 45)
$scanTabs.Size = New-Object System.Drawing.Size(1016, 292)
$scanTabs.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Left -bor
    [System.Windows.Forms.AnchorStyles]::Right
)
$scanGroup.Controls.Add($scanTabs)

$logGroup = New-Object System.Windows.Forms.GroupBox
$logGroup.Text = 'Background activity'
$logGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
$logGroup.Padding = New-Object System.Windows.Forms.Padding(12, 23, 12, 12)
$resultsSplitContainer.Panel2.Controls.Add($logGroup)

$activityLog = New-Object System.Windows.Forms.RichTextBox
$activityLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$activityLog.ReadOnly = $true
$activityLog.BackColor = [System.Drawing.Color]::FromArgb(17, 24, 39)
$activityLog.ForeColor = [System.Drawing.Color]::FromArgb(229, 231, 235)
$activityLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$activityLog.WordWrap = $false
$logGroup.Controls.Add($activityLog)

$clearLogButton = New-Object System.Windows.Forms.Button
$clearLogButton.Text = 'Clear log'
$clearLogButton.Size = New-Object System.Drawing.Size(90, 27)
$clearLogButton.Location = New-Object System.Drawing.Point(938, -2)
$clearLogButton.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Right
)
$logGroup.Controls.Add($clearLogButton)
$clearLogButton.BringToFront()

$script:ActionButtons = @($deleteServicesButton, $iisButton, $folderButton, $localAdminButton, $odbcButton, $scanButton, $whatIfButton, $selectFoldersButton)
$script:CurrentJob = $null
$script:SelectedFolderTargets = @()
$script:LatestScanResults = @()
$script:LatestVisibleScanResults = @()
$script:LatestScanInstance = ''
$script:LatestScanTimestamp = [datetime]::MinValue
$script:LatestScanHadErrors = $false
$script:LatestScanContext = $null
$script:CandidateSelectionControls = @()
$script:CandidateSelectionValid = $false
$script:ActivityLogSessionId = [guid]::NewGuid().ToString('N').Substring(0, 8)
$script:ActivityLogPath = ''
$script:ActivityLogWriteFailed = $false

function New-ActivityLogFile {
    param(
        [Parameter(Mandatory)][string]$Directory,
        [Parameter(Mandatory)][string]$SessionId
    )

    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $Directory -Force)
    }
    $safeUser = ([Environment]::UserName -replace '[^A-Za-z0-9._-]', '_')
    $fileName = 'BackgroundActivity_{0}_{1}_{2}.txt' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), $safeUser, $SessionId
    $path = Join-Path $Directory $fileName
    $header = @(
        'POST DECOM TOOL - Background Activity Log'
        'Session ID: {0}' -f $SessionId
        'Started: {0:yyyy-MM-dd HH:mm:ss zzz}' -f (Get-Date)
        'Operator: {0}\{1}' -f [Environment]::UserDomainName, [Environment]::UserName
        'Computer: {0}' -f [Environment]::MachineName
        'Tool directory: {0}' -f $PSScriptRoot
        ('-' * 100)
    )
    [System.IO.File]::WriteAllLines(
        $path,
        $header,
        (New-Object System.Text.UTF8Encoding($true))
    )
    return $path
}

function Write-ActivityLogFileLine {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Line
    )

    [System.IO.File]::AppendAllText(
        $Path,
        $Line + [Environment]::NewLine,
        (New-Object System.Text.UTF8Encoding($false))
    )
}

function Add-ActivityLog {
    param(
        [string]$Message,
        [System.Drawing.Color]$Color = ([System.Drawing.Color]::FromArgb(229, 231, 235))
    )

    if ($null -eq $Message) { return }

    $line = '[{0:HH:mm:ss}] {1}' -f (Get-Date), $Message.TrimEnd()
    $activityLog.SelectionStart = $activityLog.TextLength
    $activityLog.SelectionLength = 0
    $activityLog.SelectionColor = $Color
    $activityLog.AppendText($line + [Environment]::NewLine)
    $activityLog.SelectionColor = $activityLog.ForeColor
    $activityLog.ScrollToCaret()

    if (-not [string]::IsNullOrWhiteSpace($script:ActivityLogPath) -and
        -not $script:ActivityLogWriteFailed) {
        try {
            Write-ActivityLogFileLine -Path $script:ActivityLogPath -Line $line
        } catch {
            $script:ActivityLogWriteFailed = $true
            $warningLine = '[{0:HH:mm:ss}] WARNING: Background activity could no longer be written to the session text log: {1}' -f (Get-Date), $_.Exception.Message
            $activityLog.SelectionStart = $activityLog.TextLength
            $activityLog.SelectionColor = [System.Drawing.Color]::FromArgb(251, 191, 36)
            $activityLog.AppendText($warningLine + [Environment]::NewLine)
            $activityLog.SelectionColor = $activityLog.ForeColor
            $activityLog.ScrollToCaret()
        }
    }
}

function Get-LogColor {
    param([string]$Message)

    if ($Message -match '(?i)\b(error|failed|exit code [1-9])\b') {
        return [System.Drawing.Color]::FromArgb(248, 113, 113)
    }
    if ($Message -match '(?i)\b(warning|blocked|refusing)\b') {
        return [System.Drawing.Color]::FromArgb(251, 191, 36)
    }
    if ($Message -match '(?i)^what if:|\[whatif\]|preview') {
        return [System.Drawing.Color]::FromArgb(96, 165, 250)
    }
    if ($Message -match '(?i)successfully|verified') {
        return [System.Drawing.Color]::FromArgb(74, 222, 128)
    }
    if ($Message -match '^===') {
        return [System.Drawing.Color]::FromArgb(103, 232, 249)
    }

    return [System.Drawing.Color]::FromArgb(229, 231, 235)
}

function Set-StatusText {
    param(
        [Parameter(Mandatory)][string]$Text,
        [string]$FullText = $Text
    )

    $statusLabel.Text = $Text
    $statusToolTip.SetToolTip($statusLabel, $FullText)
}

function Set-ActionState {
    param([bool]$Busy)

    foreach ($button in $script:ActionButtons) {
        $button.Enabled = -not $Busy
    }
    $cancelButton.Enabled = $Busy
    $exportScanButton.Enabled = (-not $Busy -and @($script:LatestVisibleScanResults).Count -gt 0)
    foreach ($control in @($script:CandidateSelectionControls)) {
        if ($null -ne $control -and -not $control.IsDisposed) {
            $control.Enabled = (-not $Busy -and $script:CandidateSelectionValid)
        }
    }

    if ($Busy) {
        $progressBar.Style = 'Marquee'
        $progressBar.MarqueeAnimationSpeed = 30
    } else {
        $progressBar.MarqueeAnimationSpeed = 0
        $progressBar.Style = 'Continuous'
        $progressBar.Value = 0
    }
}

function ConvertTo-ArgumentText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    (($Text -replace '[\r\n;]+', ',') -replace ',+', ',').Trim(' ', ',')
}

function ConvertTo-ServerArray {
    param([string[]]$Text)

    @(
        $Text |
            ForEach-Object { $_ -split '[,;\s]+' } |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ } |
            Select-Object -Unique
    )
}

function Get-CombinedServerArray {
    @(
        ConvertTo-ServerArray -Text @(
            $serviceServersTextBox.Text
            $applicationServersTextBox.Text
        )
    )
}

function Get-ServiceAccountMatchDescription {
    $instance = $instanceTextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($instance)) { return '' }
    return "Account name contains: $instance"
}

function ConvertTo-EncodedFolderArgument {
    param([string[]]$FolderNames)

    @(
        $FolderNames |
            ForEach-Object {
                [System.Convert]::ToBase64String(
                    [System.Text.Encoding]::UTF8.GetBytes([string]$_)
                )
            }
    ) -join ','
}

function ConvertTo-ServiceManifestArgument {
    param([object[]]$ServiceTargets)

    $records = @(
        $ServiceTargets |
            ForEach-Object {
                $manifestTargetType = if ([string]$_.TargetType -eq 'GCS list entry') {
                    'GCSListEntry'
                } else {
                    [string]$_.TargetType
                }
                [PSCustomObject]@{
                    Server     = [string]$_.Server
                    TargetType = $manifestTargetType
                    Name       = [string]$_.Name
                    FileName   = [string]$_.FileName
                    Entry      = [string]$_.Entry
                    MatchSource = [string]$_.MatchSource
                }
            }
    )
    $json = ConvertTo-Json -InputObject $records -Compress -Depth 3
    return [System.Convert]::ToBase64String(
        [System.Text.Encoding]::UTF8.GetBytes($json)
    )
}

function ConvertTo-IISTargetManifestArgument {
    param([object[]]$IISTargets)

    $records = @(
        $IISTargets |
            ForEach-Object {
                [PSCustomObject]@{
                    Server          = [string]$_.Server
                    TargetType      = [string]$_.TargetType
                    Name            = [string]$_.Name
                    SiteName        = [string]$_.SiteName
                    ApplicationPath = [string]$_.ApplicationPath
                    MatchSource     = [string]$_.MatchSource
                }
            }
    )
    $json = ConvertTo-Json -InputObject $records -Compress -Depth 4
    return [System.Convert]::ToBase64String(
        [System.Text.Encoding]::UTF8.GetBytes($json)
    )
}

function ConvertTo-LocalAdminTargetManifestArgument {
    param([object[]]$LocalAdminTargets)

    $records = @(
        $LocalAdminTargets |
            ForEach-Object {
                [PSCustomObject]@{
                    Server = [string]$_.Server
                    Name   = [string]$_.Name
                    Path   = [string]$_.Path
                }
            }
    )
    $json = ConvertTo-Json -InputObject $records -Compress -Depth 3
    return [System.Convert]::ToBase64String(
        [System.Text.Encoding]::UTF8.GetBytes($json)
    )
}

function ConvertTo-ODBCTargetManifestArgument {
    param([object[]]$ODBCTargets)

    $records = @(
        $ODBCTargets |
            ForEach-Object {
                [PSCustomObject]@{
                    Server       = [string]$_.Server
                    Architecture = [string]$_.Architecture
                    Name         = [string]$_.Name
                }
            }
    )
    $json = ConvertTo-Json -InputObject $records -Compress -Depth 3
    return [System.Convert]::ToBase64String(
        [System.Text.Encoding]::UTF8.GetBytes($json)
    )
}

function Update-FolderSelectionSummary {
    $targets = @($script:SelectedFolderTargets)
    if ($targets.Count -eq 0) {
        $folderSelectionSummary.Text = 'No folders selected'
        return
    }

    $serverCount = @($targets | ForEach-Object { $_.Server } | Select-Object -Unique).Count
    $folderCount = @($targets | ForEach-Object { $_.Name } | Select-Object -Unique).Count
    $folderSelectionSummary.Text = '{0} server/folder target(s) selected across {1} server(s) ({2} unique folder name(s))' -f $targets.Count, $serverCount, $folderCount
}

function Clear-FolderSelection {
    param([switch]$WriteLog)

    $hadSelection = @($script:SelectedFolderTargets).Count -gt 0
    $script:SelectedFolderTargets = @()
    Update-FolderSelectionSummary

    if ($WriteLog -and $hadSelection) {
        Add-ActivityLog -Message 'Folder selection cleared because a server list changed.' -Color ([System.Drawing.Color]::FromArgb(251, 191, 36))
    }
}

function Clear-OptInCandidateSelections {
    param([switch]$WriteLog)

    $hadSelection = @(
        $script:LatestScanResults |
            Where-Object { [bool]$_.IsOptInCandidate -and [bool]$_.OptInSelected }
    ).Count -gt 0
    foreach ($record in @($script:LatestScanResults | Where-Object { [bool]$_.IsOptInCandidate })) {
        if ($null -ne $record.PSObject.Properties['OptInSelected']) {
            $record.OptInSelected = $false
        }
    }
    foreach ($control in @($script:CandidateSelectionControls)) {
        if ($control -is [System.Windows.Forms.DataGridView] -and -not $control.IsDisposed) {
            foreach ($row in $control.Rows) {
                if ($row.Cells.Count -gt 0) { $row.Cells[0].Value = $false }
            }
        }
        if ($null -ne $control -and -not $control.IsDisposed) {
            $control.Enabled = $false
        }
    }
    $script:CandidateSelectionValid = $false
    $script:LatestScanContext = $null
    if ($WriteLog -and $hadSelection) {
        Add-ActivityLog -Message 'Name-only candidate selections were cleared because an input changed. Run SCAN again before selecting candidates.' -Color ([System.Drawing.Color]::FromArgb(251, 191, 36))
    }
}

function Test-OptInCandidateSelected {
    param(
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$Name
    )

    if (-not $script:CandidateSelectionValid) { return $false }
    foreach ($record in @($script:LatestScanResults)) {
        if (
            [string]$record.Category -eq $Category -and
            [bool]$record.IsOptInCandidate -and
            [bool]$record.OptInSelected -and
            ([string]$record.Server).Equals($Server, [System.StringComparison]::OrdinalIgnoreCase) -and
            ([string]$record.Name).Equals($Name, [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            return $true
        }
    }
    return $false
}

function Get-VisibleScanResults {
    param([object[]]$ScanResults)

    @(
        $ScanResults |
            Where-Object { -not [bool]$_.IsOptInCandidate -or [bool]$_.OptInSelected }
    )
}

function Show-InputError {
    param([string]$Message)

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        'Missing or invalid input',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
}

function Add-ScanResultTab {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TabControl]$TabControl,
        [Parameter(Mandatory)][string]$Title,
        [object[]]$Records,
        [Parameter(Mandatory)][object[]]$Columns
    )

    $rows = @($Records)
    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = '{0} ({1})' -f $Title, $rows.Count
    $tab.Padding = New-Object System.Windows.Forms.Padding(6)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.AutoGenerateColumns = $false
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $true
    $grid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None
    $grid.RowTemplate.Height = 24
    $grid.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText

    foreach ($columnDefinition in $Columns) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.HeaderText = [string]$columnDefinition.Header
        $column.Name = [string]$columnDefinition.Property
        $column.Width = [int]$columnDefinition.Width
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        [void]$grid.Columns.Add($column)
    }

    foreach ($record in $rows) {
        $rowIndex = $grid.Rows.Add()
        $row = $grid.Rows[$rowIndex]
        for ($columnIndex = 0; $columnIndex -lt $Columns.Count; $columnIndex++) {
            $propertyName = [string]$Columns[$columnIndex].Property
            $value = $record.PSObject.Properties[$propertyName].Value
            $row.Cells[$columnIndex].Value = if ($null -eq $value) { '' } else { [string]$value }
        }

        if ([string]$record.Category -eq 'ApplicationPools' -and -not [bool]$record.Planned) {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(153, 27, 27)
        } elseif ([string]$record.Category -eq 'Folders' -and [bool]$record.MatchesToken) {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(220, 252, 231)
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
        }
    }

    $tab.Controls.Add($grid)
    [void]$TabControl.TabPages.Add($tab)
}

function New-ReadOnlyScanGrid {
    param(
        [object[]]$Records,
        [Parameter(Mandatory)][object[]]$Columns
    )

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.AutoGenerateColumns = $false
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $true
    $grid.RowTemplate.Height = 24
    $grid.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText

    foreach ($definition in $Columns) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.HeaderText = [string]$definition.Header
        $column.Name = [string]$definition.Property
        $column.Width = [int]$definition.Width
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        [void]$grid.Columns.Add($column)
    }
    foreach ($record in @($Records)) {
        $rowIndex = $grid.Rows.Add()
        $row = $grid.Rows[$rowIndex]
        for ($index = 0; $index -lt $Columns.Count; $index++) {
            $propertyName = [string]$Columns[$index].Property
            $property = $record.PSObject.Properties[$propertyName]
            $row.Cells[$index].Value = if ($null -eq $property -or $null -eq $property.Value) { '' } else { [string]$property.Value }
        }
        if ([string]$record.Category -eq 'ApplicationPools' -and -not [bool]$record.Planned) {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(153, 27, 27)
        }
    }
    return $grid
}

function Add-OptInCandidateScanResultTab {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TabControl]$TabControl,
        [Parameter(Mandatory)][string]$Title,
        [object[]]$Records,
        [Parameter(Mandatory)][object[]]$AutomaticColumns,
        [Parameter(Mandatory)][object[]]$CandidateColumns,
        [Parameter(Mandatory)][string]$AutomaticHeading,
        [Parameter(Mandatory)][string]$CandidateHeading,
        [Parameter(Mandatory)][string]$CandidateExplanation
    )

    $allRows = @($Records)
    $automaticRows = @($allRows | Where-Object { -not [bool]$_.IsOptInCandidate })
    $candidateRows = @($allRows | Where-Object { [bool]$_.IsOptInCandidate })
    foreach ($record in $candidateRows) {
        if ($null -eq $record.PSObject.Properties['OptInSelected']) {
            $record | Add-Member -NotePropertyName OptInSelected -NotePropertyValue $false
        }
    }

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = '{0} ({1})' -f $Title, $allRows.Count
    $tab.Padding = New-Object System.Windows.Forms.Padding(6)
    [void]$TabControl.TabPages.Add($tab)

    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $tab.Controls.Add($split)
    $split.Dock = [System.Windows.Forms.DockStyle]::Fill
    $split.Panel1MinSize = 70
    $split.Panel2MinSize = 105
    $split.SplitterDistance = 105

    $automaticLabel = New-Object System.Windows.Forms.Label
    $automaticLabel.Text = '{0} ({1})' -f $AutomaticHeading, $automaticRows.Count
    $automaticLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $automaticLabel.Height = 22
    $automaticLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
    $automaticGrid = New-ReadOnlyScanGrid -Records $automaticRows -Columns $AutomaticColumns
    $automaticGrid.Dock = [System.Windows.Forms.DockStyle]::None
    $automaticGrid.Location = New-Object System.Drawing.Point(0, 22)
    $automaticGrid.Size = New-Object System.Drawing.Size(
        $split.Panel1.ClientSize.Width,
        ([math]::Max(40, $split.Panel1.ClientSize.Height - 22))
    )
    $automaticGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $split.Panel1.Controls.Add($automaticGrid)
    $split.Panel1.Controls.Add($automaticLabel)
    $automaticLabel.BringToFront()

    $candidatePanel = New-Object System.Windows.Forms.Panel
    $candidatePanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $candidatePanel.BackColor = [System.Drawing.Color]::FromArgb(255, 247, 237)
    $split.Panel2.Controls.Add($candidatePanel)

    $candidateLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $candidateLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $candidateLayout.ColumnCount = 1
    $candidateLayout.RowCount = 2
    [void]$candidateLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$candidateLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 48)))
    [void]$candidateLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $candidateLayout.Margin = New-Object System.Windows.Forms.Padding(0)
    $candidateLayout.Padding = New-Object System.Windows.Forms.Padding(0)
    $candidatePanel.Controls.Add($candidateLayout)

    $candidateHeaderPanel = New-Object System.Windows.Forms.Panel
    $candidateHeaderPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $candidateHeaderPanel.Height = 48
    $candidateHeaderPanel.Margin = New-Object System.Windows.Forms.Padding(0)
    $candidateHeaderPanel.BackColor = $candidatePanel.BackColor

    $candidateHeader = New-Object System.Windows.Forms.Label
    $candidateHeader.Text = '{0} ({1})' -f $CandidateHeading, $candidateRows.Count
    $candidateHeader.Location = New-Object System.Drawing.Point(8, 5)
    $candidateHeader.AutoSize = $true
    $candidateHeader.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
    $candidateHeader.ForeColor = [System.Drawing.Color]::FromArgb(154, 52, 18)
    $candidateHeaderPanel.Controls.Add($candidateHeader)

    $candidateInfo = New-Object System.Windows.Forms.Label
    $candidateInfo.Text = $CandidateExplanation
    $candidateInfo.Location = New-Object System.Drawing.Point(8, 25)
    $candidateInfo.Size = New-Object System.Drawing.Size(760, 20)
    $candidateInfo.AutoEllipsis = $true
    $candidateInfo.ForeColor = [System.Drawing.Color]::FromArgb(146, 64, 14)
    $candidateInfo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $candidateHeaderPanel.Controls.Add($candidateInfo)

    $candidateButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $candidateButtonPanel.Dock = [System.Windows.Forms.DockStyle]::Right
    $candidateButtonPanel.Width = 190
    $candidateButtonPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $candidateButtonPanel.WrapContents = $false
    $candidateButtonPanel.Padding = New-Object System.Windows.Forms.Padding(4, 5, 0, 0)
    $candidateButtonPanel.BackColor = $candidatePanel.BackColor
    $candidateHeaderPanel.Controls.Add($candidateButtonPanel)
    $candidateButtonPanel.BringToFront()

    $yesAllButton = New-Object System.Windows.Forms.Button
    $yesAllButton.Text = 'Yes to all'
    $yesAllButton.Size = New-Object System.Drawing.Size(82, 25)
    $candidateButtonPanel.Controls.Add($yesAllButton)

    $noAllButton = New-Object System.Windows.Forms.Button
    $noAllButton.Text = 'No to all'
    $noAllButton.Size = New-Object System.Drawing.Size(82, 25)
    $candidateButtonPanel.Controls.Add($noAllButton)

    $candidateGrid = New-Object System.Windows.Forms.DataGridView
    $candidateGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $candidateGrid.Margin = New-Object System.Windows.Forms.Padding(0)
    $candidateGrid.ReadOnly = $false
    $candidateGrid.AllowUserToAddRows = $false
    $candidateGrid.AllowUserToDeleteRows = $false
    $candidateGrid.AllowUserToResizeRows = $false
    $candidateGrid.AutoGenerateColumns = $false
    $candidateGrid.BackgroundColor = [System.Drawing.Color]::White
    $candidateGrid.RowHeadersVisible = $false
    $candidateGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $candidateGrid.RowTemplate.Height = 24
    $candidateGrid.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter

    $checkColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $checkColumn.HeaderText = 'Decommission?'
    $checkColumn.Name = 'OptInSelected'
    $checkColumn.Width = 110
    $checkColumn.ReadOnly = $false
    [void]$candidateGrid.Columns.Add($checkColumn)
    foreach ($definition in $CandidateColumns) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.HeaderText = [string]$definition.Header
        $column.Name = [string]$definition.Property
        $column.Width = [int]$definition.Width
        $column.ReadOnly = $true
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        [void]$candidateGrid.Columns.Add($column)
    }
    foreach ($record in $candidateRows) {
        $rowIndex = $candidateGrid.Rows.Add()
        $row = $candidateGrid.Rows[$rowIndex]
        $row.Tag = $record
        $row.Cells[0].Value = [bool]$record.OptInSelected
        for ($index = 0; $index -lt $CandidateColumns.Count; $index++) {
            $propertyName = [string]$CandidateColumns[$index].Property
            $property = $record.PSObject.Properties[$propertyName]
            $row.Cells[$index + 1].Value = if ($null -eq $property -or $null -eq $property.Value) { '' } else { [string]$property.Value }
        }
    }
    $candidateGrid.Add_CurrentCellDirtyStateChanged({
        param($sender, $eventArgs)
        if ($sender.IsCurrentCellDirty) {
            [void]$sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })
    $candidateGrid.Add_CellValueChanged({
        param($sender, $eventArgs)
        if ($eventArgs.RowIndex -lt 0 -or $eventArgs.ColumnIndex -ne 0) { return }
        $row = $sender.Rows[$eventArgs.RowIndex]
        if ($null -ne $row.Tag) {
            $row.Tag.OptInSelected = [bool]$row.Cells[0].Value
        }
    })
    $candidateLayout.Controls.Add($candidateHeaderPanel, 0, 0)
    $candidateLayout.Controls.Add($candidateGrid, 0, 1)

    $yesAllButton.Tag = $candidateGrid
    $yesAllButton.Add_Click({
        param($sender, $eventArgs)
        foreach ($row in $sender.Tag.Rows) {
            $row.Cells[0].Value = $true
            if ($null -ne $row.Tag) { $row.Tag.OptInSelected = $true }
        }
    })
    $noAllButton.Tag = $candidateGrid
    $noAllButton.Add_Click({
        param($sender, $eventArgs)
        foreach ($row in $sender.Tag.Rows) {
            $row.Cells[0].Value = $false
            if ($null -ne $row.Tag) { $row.Tag.OptInSelected = $false }
        }
    })
    $yesAllButton.Enabled = $candidateRows.Count -gt 0
    $noAllButton.Enabled = $candidateRows.Count -gt 0
    $script:CandidateSelectionControls += @($candidateGrid, $yesAllButton, $noAllButton)

}

function New-NameOnlyCandidateGrid {
    param(
        [object[]]$Records,
        [Parameter(Mandatory)][object[]]$Columns
    )

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grid.Margin = New-Object System.Windows.Forms.Padding(0)
    $grid.ReadOnly = $false
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.AutoGenerateColumns = $false
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.RowTemplate.Height = 25
    $grid.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter

    $checkColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $checkColumn.HeaderText = 'Add to scan results?'
    $checkColumn.Name = 'OptInSelected'
    $checkColumn.Width = 125
    $checkColumn.ReadOnly = $false
    [void]$grid.Columns.Add($checkColumn)
    foreach ($definition in $Columns) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.HeaderText = [string]$definition.Header
        $column.Name = [string]$definition.Property
        $column.Width = [int]$definition.Width
        $column.ReadOnly = $true
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        [void]$grid.Columns.Add($column)
    }

    foreach ($record in @($Records)) {
        if ($null -eq $record.PSObject.Properties['OptInSelected']) {
            $record | Add-Member -NotePropertyName OptInSelected -NotePropertyValue $false
        }
        $rowIndex = $grid.Rows.Add()
        $row = $grid.Rows[$rowIndex]
        $row.Tag = $record
        $row.Cells[0].Value = [bool]$record.OptInSelected
        for ($index = 0; $index -lt $Columns.Count; $index++) {
            $propertyName = [string]$Columns[$index].Property
            $property = $record.PSObject.Properties[$propertyName]
            $row.Cells[$index + 1].Value = if ($null -eq $property -or $null -eq $property.Value) { '' } else { [string]$property.Value }
        }
    }
    $grid.Add_CurrentCellDirtyStateChanged({
        param($sender, $eventArgs)
        if ($sender.IsCurrentCellDirty) {
            [void]$sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })
    $grid.Add_CellValueChanged({
        param($sender, $eventArgs)
        if ($eventArgs.RowIndex -lt 0 -or $eventArgs.ColumnIndex -ne 0) { return }
        $row = $sender.Rows[$eventArgs.RowIndex]
        if ($null -ne $row.Tag) {
            $row.Tag.OptInSelected = [bool]$row.Cells[0].Value
        }
    })
    return $grid
}

function Show-NameOnlyCandidateSelectionDialog {
    param(
        [object[]]$ScanResults,
        [string]$Instance,
        [string]$RenderPreviewPath
    )

    $serviceCandidates = @($ScanResults | Where-Object { $_.Category -eq 'Services' -and [bool]$_.IsOptInCandidate })
    $poolCandidates = @($ScanResults | Where-Object { $_.Category -eq 'ApplicationPools' -and [bool]$_.IsOptInCandidate })
    if ($serviceCandidates.Count -eq 0 -and $poolCandidates.Count -eq 0) {
        return $true
    }

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = 'Review name-only candidates'
    $dialog.StartPosition = 'CenterParent'
    $dialog.Size = New-Object System.Drawing.Size(1250, 680)
    $dialog.MinimumSize = New-Object System.Drawing.Size(900, 520)
    $dialog.BackColor = [System.Drawing.Color]::FromArgb(255, 247, 237)
    $dialog.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $dialog.ShowInTaskbar = $false

    $heading = New-Object System.Windows.Forms.Label
    $heading.Text = 'Review name-only service and application-pool candidates'
    $heading.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 15)
    $heading.ForeColor = [System.Drawing.Color]::FromArgb(154, 52, 18)
    $heading.Location = New-Object System.Drawing.Point(20, 14)
    $heading.AutoSize = $true
    $dialog.Controls.Add($heading)

    $instructions = New-Object System.Windows.Forms.Label
    $instructions.Text = "Instance token: $Instance`r`nThese names contain the token, but their service Log On As or application-pool identity does not. Check only the rows to add to Scan Results and the later decommission confirmation."
    $instructions.Location = New-Object System.Drawing.Point(22, 50)
    $instructions.Size = New-Object System.Drawing.Size(1190, 44)
    $instructions.ForeColor = [System.Drawing.Color]::FromArgb(146, 64, 14)
    $instructions.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $dialog.Controls.Add($instructions)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(20, 100)
    $tabs.Size = New-Object System.Drawing.Size(1194, 475)
    $tabs.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $dialog.Controls.Add($tabs)

    $candidateGrids = New-Object System.Collections.ArrayList
    $serviceTab = New-Object System.Windows.Forms.TabPage
    $serviceTab.Text = 'Services ({0})' -f $serviceCandidates.Count
    $serviceTab.Padding = New-Object System.Windows.Forms.Padding(5)
    $serviceGrid = New-NameOnlyCandidateGrid -Records $serviceCandidates -Columns @(
        [PSCustomObject]@{ Header = 'Server'; Property = 'Server'; Width = 155 }
        [PSCustomObject]@{ Header = 'Service name'; Property = 'Name'; Width = 240 }
        [PSCustomObject]@{ Header = 'Display name'; Property = 'DisplayName'; Width = 240 }
        [PSCustomObject]@{ Header = 'State'; Property = 'State'; Width = 90 }
        [PSCustomObject]@{ Header = 'Log On As'; Property = 'StartName'; Width = 230 }
        [PSCustomObject]@{ Header = 'Reason'; Property = 'MatchReason'; Width = 380 }
    )
    $serviceTab.Controls.Add($serviceGrid)
    [void]$tabs.TabPages.Add($serviceTab)
    [void]$candidateGrids.Add($serviceGrid)

    $poolTab = New-Object System.Windows.Forms.TabPage
    $poolTab.Text = 'Application Pools ({0})' -f $poolCandidates.Count
    $poolTab.Padding = New-Object System.Windows.Forms.Padding(5)
    $poolGrid = New-NameOnlyCandidateGrid -Records $poolCandidates -Columns @(
        [PSCustomObject]@{ Header = 'Server'; Property = 'Server'; Width = 155 }
        [PSCustomObject]@{ Header = 'Application pool'; Property = 'Name'; Width = 270 }
        [PSCustomObject]@{ Header = 'State'; Property = 'State'; Width = 90 }
        [PSCustomObject]@{ Header = 'Pool identity'; Property = 'PoolIdentity'; Width = 250 }
        [PSCustomObject]@{ Header = 'Removal status'; Property = 'RemovalStatus'; Width = 125 }
        [PSCustomObject]@{ Header = 'Reason / dependency'; Property = 'ScanDetails'; Width = 430 }
    )
    $poolTab.Controls.Add($poolGrid)
    [void]$tabs.TabPages.Add($poolTab)
    [void]$candidateGrids.Add($poolGrid)

    $selectAllButton = New-Object System.Windows.Forms.Button
    $selectAllButton.Text = 'Check all on tab'
    $selectAllButton.Location = New-Object System.Drawing.Point(20, 590)
    $selectAllButton.Size = New-Object System.Drawing.Size(130, 32)
    $selectAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $dialog.Controls.Add($selectAllButton)

    $clearAllButton = New-Object System.Windows.Forms.Button
    $clearAllButton.Text = 'Clear all on tab'
    $clearAllButton.Location = New-Object System.Drawing.Point(158, 590)
    $clearAllButton.Size = New-Object System.Drawing.Size(130, 32)
    $clearAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $dialog.Controls.Add($clearAllButton)

    $selectAllButton.Tag = $tabs
    $selectAllButton.Add_Click({
        param($sender, $eventArgs)
        $activeGrid = @($sender.Tag.SelectedTab.Controls | Where-Object { $_ -is [System.Windows.Forms.DataGridView] }) | Select-Object -First 1
        foreach ($row in @($activeGrid.Rows)) {
            $row.Cells[0].Value = $true
            if ($null -ne $row.Tag) { $row.Tag.OptInSelected = $true }
        }
    })
    $clearAllButton.Tag = $tabs
    $clearAllButton.Add_Click({
        param($sender, $eventArgs)
        $activeGrid = @($sender.Tag.SelectedTab.Controls | Where-Object { $_ -is [System.Windows.Forms.DataGridView] }) | Select-Object -First 1
        foreach ($row in @($activeGrid.Rows)) {
            $row.Cells[0].Value = $false
            if ($null -ne $row.Tag) { $row.Tag.OptInSelected = $false }
        }
    })

    $cancelDialogButton = New-Object System.Windows.Forms.Button
    $cancelDialogButton.Text = 'Cancel'
    $cancelDialogButton.Location = New-Object System.Drawing.Point(900, 590)
    $cancelDialogButton.Size = New-Object System.Drawing.Size(105, 32)
    $cancelDialogButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $cancelDialogButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialog.CancelButton = $cancelDialogButton
    $dialog.Controls.Add($cancelDialogButton)

    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Text = 'Add Selected to Scan Results'
    $addButton.Location = New-Object System.Drawing.Point(1014, 590)
    $addButton.Size = New-Object System.Drawing.Size(200, 32)
    $addButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $addButton.BackColor = [System.Drawing.Color]::FromArgb(217, 119, 6)
    $addButton.ForeColor = [System.Drawing.Color]::White
    $addButton.FlatStyle = 'Flat'
    $addButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dialog.AcceptButton = $addButton
    $dialog.Controls.Add($addButton)

    if (-not [string]::IsNullOrWhiteSpace($RenderPreviewPath)) {
        $dialog.Show()
        [System.Windows.Forms.Application]::DoEvents()
        $bitmap = New-Object System.Drawing.Bitmap($dialog.Width, $dialog.Height)
        try {
            $dialog.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, $dialog.Width, $dialog.Height)))
            $bitmap.Save($RenderPreviewPath, [System.Drawing.Imaging.ImageFormat]::Png)
        } finally {
            $bitmap.Dispose()
            $dialog.Close()
            $dialog.Dispose()
        }
        return $false
    }

    try {
        $result = $dialog.ShowDialog($form)
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            foreach ($candidate in @($serviceCandidates + $poolCandidates)) {
                $candidate.OptInSelected = $false
            }
            return $false
        }
        foreach ($grid in @($candidateGrids)) {
            if ($grid.IsCurrentCellDirty) {
                [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
            }
        }
        return $true
    } finally {
        $dialog.Dispose()
    }
}

function ConvertTo-ScanExportRows {
    param(
        [object[]]$ScanResults,
        [string]$Instance,
        [datetime]$ScanTimestamp = [datetime]::MinValue,
        [bool]$HadDiscoveryErrors = $false
    )

    $timestampText = if ($ScanTimestamp -eq [datetime]::MinValue) {
        ''
    } else {
        $ScanTimestamp.ToString('yyyy-MM-dd HH:mm:ss')
    }

    @(
        foreach ($record in @($ScanResults)) {
            $values = @{}
            foreach ($property in $record.PSObject.Properties) {
                $values[[string]$property.Name] = $property.Value
            }

            [PSCustomObject]@{
                ScanTimestamp     = $timestampText
                InstanceToken     = $Instance
                ScanIncomplete    = $HadDiscoveryErrors
                Category          = [string]$values['Category']
                Server            = [string]$values['Server']
                TargetType        = [string]$values['TargetType']
                Name              = [string]$values['Name']
                DisplayName       = [string]$values['DisplayName']
                State             = [string]$values['State']
                LogOnAs           = [string]$values['StartName']
                FileName          = [string]$values['FileName']
                Entry             = [string]$values['Entry']
                SiteName          = [string]$values['SiteName']
                ApplicationPath   = [string]$values['ApplicationPath']
                AppPoolName       = [string]$values['AppPoolName']
                Architecture      = [string]$values['Architecture']
                Driver            = [string]$values['Driver']
                KeyDriver         = [string]$values['KeyDriver']
                DataServer        = [string]$values['DataServer']
                Database          = [string]$values['Database']
                Description       = [string]$values['Description']
                Path              = [string]$values['Path']
                RegistryPath      = [string]$values['RegistryPath']
                Details           = [string]$values['Details']
                BlockedReason     = [string]$values['BlockedReason']
                EligibleForRemoval = [bool]$values['Planned']
                RemovalStatus     = [string]$values['RemovalStatus']
                MatchesToken      = [bool]$values['MatchesToken']
                MatchSource       = [string]$values['MatchSource']
                MatchReason       = [string]$values['MatchReason']
                PoolIdentity      = [string]$values['PoolIdentity']
                IsOptInCandidate  = [bool]$values['IsOptInCandidate']
                IncludedByOperator = [bool]$values['OptInSelected']
            }
        }
    )
}

function ConvertTo-ExcelColumnName {
    param([Parameter(Mandatory)][int]$ColumnNumber)

    $name = ''
    $value = $ColumnNumber
    while ($value -gt 0) {
        $value--
        $name = [char](65 + ($value % 26)) + $name
        $value = [math]::Floor($value / 26)
    }
    return $name
}

function ConvertTo-OpenXmlText {
    param($Value)

    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    $text = [regex]::Replace($text, '[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]', '')
    return [System.Security.SecurityElement]::Escape($text)
}

function Add-OpenXmlZipEntry {
    param(
        [Parameter(Mandatory)]$Archive,
        [Parameter(Mandatory)][string]$EntryName,
        [Parameter(Mandatory)][string]$Content
    )

    $entry = $Archive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $writer = New-Object System.IO.StreamWriter(
        $stream,
        (New-Object System.Text.UTF8Encoding($false))
    )
    try {
        $writer.Write($Content)
    } finally {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function Get-ScanWorkbookSheetDefinitions {
    param(
        [object[]]$ScanResults,
        [string]$Instance,
        [datetime]$ScanTimestamp = [datetime]::MinValue
    )

    $timestampText = if ($ScanTimestamp -eq [datetime]::MinValue) {
        ''
    } else {
        $ScanTimestamp.ToString('yyyy-MM-dd HH:mm:ss')
    }
    $results = @($ScanResults)

    $definitions = @(
        [PSCustomObject]@{
            Name = 'Services'
            Headers = @('Scan timestamp', 'Instance token', 'Server', 'Type', 'Service / list entry', 'Display name / file', 'State', 'Log On As', 'Match type', 'Path', 'Match reason')
            Widths = @(20, 18, 20, 20, 30, 30, 14, 32, 22, 52, 55)
            Rows = @(
                foreach ($record in @($results | Where-Object { $_.Category -eq 'Services' })) {
                    [PSCustomObject]@{ Values = @($timestampText, $Instance, [string]$record.Server, [string]$record.TargetType, [string]$record.Name, [string]$record.DisplayName, [string]$record.State, [string]$record.StartName, [string]$record.SelectionText, [string]$record.Path, [string]$record.MatchReason) }
                }
            )
        }
        [PSCustomObject]@{
            Name = 'IIS Sites'
            Headers = @('Scan timestamp', 'Instance token', 'Server', 'Site', 'State', 'Details')
            Widths = @(20, 18, 20, 34, 14, 70)
            Rows = @(
                foreach ($record in @($results | Where-Object { $_.Category -eq 'IISSites' })) {
                    [PSCustomObject]@{ Values = @($timestampText, $Instance, [string]$record.Server, [string]$record.Name, [string]$record.State, [string]$record.Details) }
                }
            )
        }
        [PSCustomObject]@{
            Name = 'Application Pools'
            Headers = @('Scan timestamp', 'Instance token', 'Server', 'Application pool', 'State', 'Pool identity', 'Match type', 'Removal status', 'Details', 'Blocked reason')
            Widths = @(20, 18, 20, 36, 14, 32, 22, 18, 60, 60)
            Rows = @(
                foreach ($record in @($results | Where-Object { $_.Category -eq 'ApplicationPools' })) {
                    [PSCustomObject]@{ Values = @($timestampText, $Instance, [string]$record.Server, [string]$record.Name, [string]$record.State, [string]$record.PoolIdentity, [string]$record.SelectionText, [string]$record.RemovalStatus, [string]$record.Details, [string]$record.BlockedReason) }
                }
            )
        }
        [PSCustomObject]@{
            Name = 'ODBC Data Sources'
            Headers = @('Scan timestamp', 'Instance token', 'Server', 'Architecture', 'System DSN', 'Driver', 'Data server', 'Database', 'Registry path')
            Widths = @(20, 18, 20, 15, 30, 35, 28, 25, 65)
            Rows = @(
                foreach ($record in @($results | Where-Object { $_.Category -eq 'ODBCDataSources' })) {
                    [PSCustomObject]@{ Values = @($timestampText, $Instance, [string]$record.Server, [string]$record.Architecture, [string]$record.Name, [string]$record.Driver, [string]$record.DataServer, [string]$record.Database, [string]$record.RegistryPath) }
                }
            )
        }
        [PSCustomObject]@{
            Name = 'Local Admins'
            Headers = @('Scan timestamp', 'Instance token', 'Server', 'Local group', 'Matching principal', 'Class', 'Principal path')
            Widths = @(20, 18, 20, 24, 36, 16, 65)
            Rows = @(
                foreach ($record in @($results | Where-Object { $_.Category -eq 'LocalAdmins' })) {
                    [PSCustomObject]@{ Values = @($timestampText, $Instance, [string]$record.Server, [string]$record.DisplayName, [string]$record.Name, [string]$record.TargetType, [string]$record.Path) }
                }
            )
        }
        [PSCustomObject]@{
            Name = 'Client Folders'
            Headers = @('Scan timestamp', 'Instance token', 'Server', 'Folder', 'Token match', 'Full UNC path')
            Widths = @(20, 18, 20, 38, 16, 70)
            Rows = @(
                foreach ($record in @($results | Where-Object { $_.Category -eq 'Folders' })) {
                    [PSCustomObject]@{ Values = @($timestampText, $Instance, [string]$record.Server, [string]$record.Name, [string]$record.MatchText, [string]$record.Path) }
                }
            )
        }
    )
    return $definitions
}

function ConvertTo-OpenXmlWorksheetXml {
    param(
        [Parameter(Mandatory)]$Definition
    )

    $headers = @($Definition.Headers)
    $rows = @($Definition.Rows)
    $lastColumn = ConvertTo-ExcelColumnName -ColumnNumber $headers.Count
    $lastRow = [math]::Max(1, $rows.Count + 1)
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$builder.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$builder.Append('<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    [void]$builder.Append('<cols>')
    for ($index = 0; $index -lt $headers.Count; $index++) {
        $columnNumber = $index + 1
        $width = [double]$Definition.Widths[$index]
        [void]$builder.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f $columnNumber, $width.ToString([System.Globalization.CultureInfo]::InvariantCulture)))
    }
    [void]$builder.Append('</cols><sheetData>')
    [void]$builder.Append('<row r="1" ht="24" customHeight="1">')
    for ($index = 0; $index -lt $headers.Count; $index++) {
        $cellReference = '{0}1' -f (ConvertTo-ExcelColumnName -ColumnNumber ($index + 1))
        $text = ConvertTo-OpenXmlText -Value $headers[$index]
        [void]$builder.Append(('<c r="{0}" s="1" t="inlineStr"><is><t>{1}</t></is></c>' -f $cellReference, $text))
    }
    [void]$builder.Append('</row>')
    for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
        $excelRow = $rowIndex + 2
        [void]$builder.Append(('<row r="{0}">' -f $excelRow))
        $rowValues = @($rows[$rowIndex].Values)
        for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
            $cellReference = '{0}{1}' -f (ConvertTo-ExcelColumnName -ColumnNumber ($columnIndex + 1)), $excelRow
            $text = ConvertTo-OpenXmlText -Value $rowValues[$columnIndex]
            [void]$builder.Append(('<c r="{0}" t="inlineStr"><is><t xml:space="preserve">{1}</t></is></c>' -f $cellReference, $text))
        }
        [void]$builder.Append('</row>')
    }
    [void]$builder.Append('</sheetData>')
    [void]$builder.Append(('<autoFilter ref="A1:{0}{1}"/>' -f $lastColumn, $lastRow))
    [void]$builder.Append('</worksheet>')
    return $builder.ToString()
}

function Export-ScanResultsWorkbook {
    param(
        [Parameter(Mandatory)][string]$Path,
        [object[]]$ScanResults,
        [string]$Instance,
        [datetime]$ScanTimestamp = [datetime]::MinValue
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $definitions = @(Get-ScanWorkbookSheetDefinitions -ScanResults $ScanResults -Instance $Instance -ScanTimestamp $ScanTimestamp)
    $destinationPath = [System.IO.Path]::GetFullPath($Path)
    $destinationDirectory = [System.IO.Path]::GetDirectoryName($destinationPath)
    if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $destinationDirectory -Force)
    }
    $temporaryPath = Join-Path $destinationDirectory ('.{0}.{1}.tmp' -f [System.IO.Path]::GetFileName($destinationPath), [guid]::NewGuid().ToString('N'))
    $fileStream = $null
    $archive = $null
    try {
        $fileStream = [System.IO.File]::Open($temporaryPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $archive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)

        $contentTypeOverrides = New-Object System.Text.StringBuilder
        $workbookSheets = New-Object System.Text.StringBuilder
        $workbookRelationships = New-Object System.Text.StringBuilder
        for ($index = 0; $index -lt $definitions.Count; $index++) {
            $sheetNumber = $index + 1
            [void]$contentTypeOverrides.Append(('<Override PartName="/xl/worksheets/sheet{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' -f $sheetNumber))
            [void]$workbookSheets.Append(('<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>' -f (ConvertTo-OpenXmlText -Value $definitions[$index].Name), $sheetNumber))
            [void]$workbookRelationships.Append(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>' -f $sheetNumber))
            Add-OpenXmlZipEntry -Archive $archive -EntryName ('xl/worksheets/sheet{0}.xml' -f $sheetNumber) -Content (ConvertTo-OpenXmlWorksheetXml -Definition $definitions[$index])
        }
        $styleRelationshipId = $definitions.Count + 1
        [void]$workbookRelationships.Append(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' -f $styleRelationshipId))

        $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>{0}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>' -f $contentTypeOverrides.ToString()
        Add-OpenXmlZipEntry -Archive $archive -EntryName '[Content_Types].xml' -Content $contentTypes
        Add-OpenXmlZipEntry -Archive $archive -EntryName '_rels/.rels' -Content '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'
        Add-OpenXmlZipEntry -Archive $archive -EntryName 'xl/workbook.xml' -Content ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>{0}</sheets></workbook>' -f $workbookSheets.ToString())
        Add-OpenXmlZipEntry -Archive $archive -EntryName 'xl/_rels/workbook.xml.rels' -Content ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{0}</Relationships>' -f $workbookRelationships.ToString())
        Add-OpenXmlZipEntry -Archive $archive -EntryName 'xl/styles.xml' -Content '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Segoe UI"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Segoe UI"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD9E2F3"/></left><right style="thin"><color rgb="FFD9E2F3"/></right><top style="thin"><color rgb="FFD9E2F3"/></top><bottom style="thin"><color rgb="FFD9E2F3"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
        $created = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        Add-OpenXmlZipEntry -Archive $archive -EntryName 'docProps/core.xml' -Content ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>POST DECOM Scan Results</dc:title><dc:creator>{0}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">{1}</dcterms:created></cp:coreProperties>' -f (ConvertTo-OpenXmlText -Value [Environment]::UserName), $created)
        Add-OpenXmlZipEntry -Archive $archive -EntryName 'docProps/app.xml' -Content '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>POST DECOM TOOL</Application></Properties>'
    } finally {
        if ($null -ne $archive) { $archive.Dispose() }
        if ($null -ne $fileStream) { $fileStream.Dispose() }
    }

    try {
        if (Test-Path -LiteralPath $destinationPath -PathType Leaf) {
            [System.IO.File]::Delete($destinationPath)
        }
        [System.IO.File]::Move($temporaryPath, $destinationPath)
    } catch {
        Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        throw
    }
    return [PSCustomObject]@{
        Path       = $destinationPath
        SheetCount = $definitions.Count
        RowCount   = @($ScanResults).Count
    }
}

function Update-ScanResultsPanel {
    param(
        [object[]]$ScanResults,
        [string]$Instance,
        [bool]$HadDiscoveryErrors,
        [string]$RenderPreviewPath
    )

    $results = @($ScanResults)
    foreach ($record in $results) {
        if ($null -eq $record.PSObject.Properties['SelectionText']) {
            $record | Add-Member -NotePropertyName SelectionText -NotePropertyValue ''
        }
        $record.SelectionText = if ([bool]$record.IsOptInCandidate) {
            if ([bool]$record.OptInSelected) { 'Name-only (selected)' } else { 'Name-only (not selected)' }
        } else {
            'Automatic'
        }
    }
    $displayResults = @(Get-VisibleScanResults -ScanResults $results)
    $script:LatestScanResults = @($results)
    $script:LatestVisibleScanResults = @($displayResults)
    $script:LatestScanInstance = $Instance
    $script:LatestScanTimestamp = if ($results.Count -gt 0) { Get-Date } else { [datetime]::MinValue }
    $script:LatestScanHadErrors = $HadDiscoveryErrors
    $script:CandidateSelectionValid = ($results.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($Instance))
    $exportScanButton.Enabled = ($displayResults.Count -gt 0 -and $null -eq $script:CurrentJob)
    $script:CandidateSelectionControls = @()
    $scanTabs.TabPages.Clear()
    $scanSummary.Text = if ([string]::IsNullOrWhiteSpace($Instance)) {
        'Click SCAN to display matching resources. Scanning does not change configuration.'
    } else {
        "Instance token: $Instance | $($displayResults.Count) result(s) | Read-only scan - nothing was changed."
    }
    if ($HadDiscoveryErrors) {
        $scanSummary.Text += ' Some servers could not be fully scanned; review the background log.'
        $scanSummary.ForeColor = [System.Drawing.Color]::FromArgb(180, 83, 9)
    } else {
        $scanSummary.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    }

    Add-ScanResultTab -TabControl $scanTabs -Title 'Services' `
        -Records @($displayResults | Where-Object { $_.Category -eq 'Services' }) `
        -Columns @(
            [PSCustomObject]@{ Header = 'Server'; Property = 'Server'; Width = 145 }
            [PSCustomObject]@{ Header = 'Type'; Property = 'TargetType'; Width = 105 }
            [PSCustomObject]@{ Header = 'Service / list entry'; Property = 'Name'; Width = 230 }
            [PSCustomObject]@{ Header = 'Display name / file'; Property = 'DisplayName'; Width = 220 }
            [PSCustomObject]@{ Header = 'State'; Property = 'State'; Width = 80 }
            [PSCustomObject]@{ Header = 'Log On As'; Property = 'StartName'; Width = 220 }
            [PSCustomObject]@{ Header = 'Match type'; Property = 'SelectionText'; Width = 160 }
            [PSCustomObject]@{ Header = 'Path'; Property = 'Path'; Width = 330 }
        )

    Add-ScanResultTab -TabControl $scanTabs -Title 'IIS Sites' `
        -Records @($displayResults | Where-Object { $_.Category -eq 'IISSites' }) `
        -Columns @(
            [PSCustomObject]@{ Header = 'Server'; Property = 'Server'; Width = 160 }
            [PSCustomObject]@{ Header = 'Site'; Property = 'Name'; Width = 260 }
            [PSCustomObject]@{ Header = 'State'; Property = 'State'; Width = 100 }
            [PSCustomObject]@{ Header = 'Details'; Property = 'Details'; Width = 600 }
        )

    Add-ScanResultTab -TabControl $scanTabs -Title 'Application Pools' `
        -Records @($displayResults | Where-Object { $_.Category -eq 'ApplicationPools' }) `
        -Columns @(
            [PSCustomObject]@{ Header = 'Server'; Property = 'Server'; Width = 150 }
            [PSCustomObject]@{ Header = 'Application pool'; Property = 'Name'; Width = 260 }
            [PSCustomObject]@{ Header = 'State'; Property = 'State'; Width = 100 }
            [PSCustomObject]@{ Header = 'Pool identity'; Property = 'PoolIdentity'; Width = 220 }
            [PSCustomObject]@{ Header = 'Match type'; Property = 'SelectionText'; Width = 160 }
            [PSCustomObject]@{ Header = 'Removal status'; Property = 'RemovalStatus'; Width = 130 }
            [PSCustomObject]@{ Header = 'Details / blocked reason'; Property = 'ScanDetails'; Width = 520 }
        )

    Add-ScanResultTab -TabControl $scanTabs -Title 'ODBC Data Sources' `
        -Records @($displayResults | Where-Object { $_.Category -eq 'ODBCDataSources' }) `
        -Columns @(
            [PSCustomObject]@{ Header = 'Server'; Property = 'Server'; Width = 145 }
            [PSCustomObject]@{ Header = 'Architecture'; Property = 'Architecture'; Width = 100 }
            [PSCustomObject]@{ Header = 'System DSN'; Property = 'Name'; Width = 220 }
            [PSCustomObject]@{ Header = 'Driver'; Property = 'Driver'; Width = 230 }
            [PSCustomObject]@{ Header = 'Data server'; Property = 'DataServer'; Width = 180 }
            [PSCustomObject]@{ Header = 'Database'; Property = 'Database'; Width = 160 }
            [PSCustomObject]@{ Header = 'Registry path'; Property = 'RegistryPath'; Width = 380 }
        )

    Add-ScanResultTab -TabControl $scanTabs -Title 'Local Admins' `
        -Records @($displayResults | Where-Object { $_.Category -eq 'LocalAdmins' }) `
        -Columns @(
            [PSCustomObject]@{ Header = 'Server'; Property = 'Server'; Width = 160 }
            [PSCustomObject]@{ Header = 'Local group'; Property = 'DisplayName'; Width = 200 }
            [PSCustomObject]@{ Header = 'Matching principal'; Property = 'Name'; Width = 260 }
            [PSCustomObject]@{ Header = 'Class'; Property = 'TargetType'; Width = 110 }
            [PSCustomObject]@{ Header = 'Principal path'; Property = 'Path'; Width = 480 }
        )

    Add-ScanResultTab -TabControl $scanTabs -Title 'Client Folders' `
        -Records @($displayResults | Where-Object { $_.Category -eq 'Folders' }) `
        -Columns @(
            [PSCustomObject]@{ Header = 'Server'; Property = 'Server'; Width = 180 }
            [PSCustomObject]@{ Header = 'Folder'; Property = 'Name'; Width = 300 }
            [PSCustomObject]@{ Header = 'Token match'; Property = 'MatchText'; Width = 110 }
            [PSCustomObject]@{ Header = 'Full UNC path'; Property = 'Path'; Width = 550 }
        )

    if (-not [string]::IsNullOrWhiteSpace($RenderPreviewPath)) {
        if ($RenderScanResultsWidePreview) {
            $form.Size = New-Object System.Drawing.Size(1900, 1000)
        }
        if (-not [string]::IsNullOrWhiteSpace($RenderScanResultsPreviewTab)) {
            foreach ($previewTab in $scanTabs.TabPages) {
                if ($previewTab.Text.StartsWith($RenderScanResultsPreviewTab, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $scanTabs.SelectedTab = $previewTab
                    break
                }
            }
        }
        $form.Show()
        [System.Windows.Forms.Application]::DoEvents()
        $bitmap = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
        try {
            $form.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, $form.Width, $form.Height)))
            $bitmap.Save($RenderPreviewPath, [System.Drawing.Imaging.ImageFormat]::Png)
        } finally {
            $bitmap.Dispose()
            $form.Hide()
        }
    }
}

function Show-ServiceDeletionConfirmation {
    param(
        [object[]]$ServiceTargets,
        [string]$AccountToken,
        [string]$RenderPreviewPath
    )

    $confirmation = New-Object System.Windows.Forms.Form
    $confirmation.Text = 'Delete services'
    $confirmation.StartPosition = 'CenterParent'
    $confirmation.Size = New-Object System.Drawing.Size(1120, 600)
    $confirmation.MinimumSize = New-Object System.Drawing.Size(820, 500)
    $confirmation.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $confirmation.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $confirmation.ShowInTaskbar = $false

    $heading = New-Object System.Windows.Forms.Label
    $heading.Text = 'Confirm service and GCS list cleanup targets'
    $heading.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $heading.ForeColor = [System.Drawing.Color]::FromArgb(153, 27, 27)
    $heading.Location = New-Object System.Drawing.Point(20, 14)
    $heading.AutoSize = $true
    $confirmation.Controls.Add($heading)

    $details = New-Object System.Windows.Forms.Label
    $details.Text = "Instance token: $AccountToken`r`nServices will be deleted and the listed lines will be removed from the two GCS service-list files. Every target is revalidated before modification."
    $details.Location = New-Object System.Drawing.Point(22, 48)
    $details.Size = New-Object System.Drawing.Size(1060, 38)
    $details.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    $confirmation.Controls.Add($details)

    $serviceList = New-Object System.Windows.Forms.ListView
    $serviceList.Location = New-Object System.Drawing.Point(20, 92)
    $serviceList.Size = New-Object System.Drawing.Size(1062, 410)
    $serviceList.View = [System.Windows.Forms.View]::Details
    $serviceList.FullRowSelect = $true
    $serviceList.GridLines = $true
    $serviceList.HideSelection = $false
    $serviceList.ShowItemToolTips = $true
    $serviceList.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    [void]$serviceList.Columns.Add('Server', 135)
    [void]$serviceList.Columns.Add('Target type', 105)
    [void]$serviceList.Columns.Add('Service / list entry', 245)
    [void]$serviceList.Columns.Add('Display name / file', 220)
    [void]$serviceList.Columns.Add('State', 85)
    [void]$serviceList.Columns.Add('Log On As / full path', 330)

    foreach ($serviceTarget in @($ServiceTargets | Sort-Object Server, DisplayName, Name)) {
        $targetType = if ([string]::IsNullOrWhiteSpace([string]$serviceTarget.TargetType)) {
            'Service'
        } else {
            [string]$serviceTarget.TargetType
        }
        $targetTypeDisplay = if ($targetType -eq 'GCSListEntry') {
            'GCS list entry'
        } elseif ([string]$serviceTarget.MatchSource -eq 'NameOnly') {
            'Service (name-only opt-in)'
        } else {
            $targetType
        }
        $targetName = if ($targetType -eq 'GCSListEntry') {
            [string]$serviceTarget.Entry
        } else {
            [string]$serviceTarget.Name
        }
        $sourceName = if ($targetType -eq 'GCSListEntry') {
            [string]$serviceTarget.FileName
        } else {
            [string]$serviceTarget.DisplayName
        }
        $detailsText = if ($targetType -eq 'GCSListEntry') {
            [string]$serviceTarget.Path
        } else {
            [string]$serviceTarget.StartName
        }
        $item = New-Object System.Windows.Forms.ListViewItem([string]$serviceTarget.Server)
        [void]$item.SubItems.Add($targetTypeDisplay)
        [void]$item.SubItems.Add($targetName)
        [void]$item.SubItems.Add($sourceName)
        [void]$item.SubItems.Add([string]$serviceTarget.State)
        [void]$item.SubItems.Add($detailsText)
        $item.ToolTipText = if ($targetType -eq 'GCSListEntry') {
            '{0} - remove line: {1}' -f $serviceTarget.Path, $serviceTarget.Entry
        } else {
            '{0}\{1} - {2}' -f $serviceTarget.Server, $serviceTarget.Name, $serviceTarget.DisplayName
        }
        [void]$serviceList.Items.Add($item)
    }
    $confirmation.Controls.Add($serviceList)

    $countLabel = New-Object System.Windows.Forms.Label
    $serviceCount = @($ServiceTargets | Where-Object { [string]$_.TargetType -ne 'GCSListEntry' }).Count
    $listEntryCount = @($ServiceTargets | Where-Object { [string]$_.TargetType -eq 'GCSListEntry' }).Count
    $countLabel.Text = '{0} service(s) and {1} GCS list entr{2} will be targeted.' -f $serviceCount, $listEntryCount, $(if ($listEntryCount -eq 1) { 'y' } else { 'ies' })
    $countLabel.Location = New-Object System.Drawing.Point(22, 512)
    $countLabel.Size = New-Object System.Drawing.Size(600, 30)
    $countLabel.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left
    )
    $countLabel.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    $confirmation.Controls.Add($countLabel)

    $cancelServiceButton = New-Object System.Windows.Forms.Button
    $cancelServiceButton.Text = 'Cancel'
    $cancelServiceButton.Location = New-Object System.Drawing.Point(828, 520)
    $cancelServiceButton.Size = New-Object System.Drawing.Size(105, 32)
    $cancelServiceButton.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $cancelServiceButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $confirmation.CancelButton = $cancelServiceButton
    $confirmation.Controls.Add($cancelServiceButton)

    $deleteServiceButton = New-Object System.Windows.Forms.Button
    $deleteServiceButton.Text = 'Delete Listed Targets'
    $deleteServiceButton.Location = New-Object System.Drawing.Point(940, 520)
    $deleteServiceButton.Size = New-Object System.Drawing.Size(142, 32)
    $deleteServiceButton.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $deleteServiceButton.BackColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
    $deleteServiceButton.ForeColor = [System.Drawing.Color]::White
    $deleteServiceButton.FlatStyle = 'Flat'
    $deleteServiceButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $confirmation.Controls.Add($deleteServiceButton)

    if (-not [string]::IsNullOrWhiteSpace($RenderPreviewPath)) {
        $confirmation.Show()
        [System.Windows.Forms.Application]::DoEvents()
        $bitmap = New-Object System.Drawing.Bitmap($confirmation.Width, $confirmation.Height)
        try {
            $confirmation.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, $confirmation.Width, $confirmation.Height)))
            $bitmap.Save($RenderPreviewPath, [System.Drawing.Imaging.ImageFormat]::Png)
        } finally {
            $bitmap.Dispose()
            $confirmation.Close()
            $confirmation.Dispose()
        }
        return $false
    }

    try {
        return ($confirmation.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::Yes)
    } finally {
        $confirmation.Dispose()
    }
}

function Show-LocalAdminCleanupConfirmation {
    param(
        [object[]]$LocalAdminTargets,
        [string]$Instance,
        [string]$RenderPreviewPath
    )

    $confirmation = New-Object System.Windows.Forms.Form
    $confirmation.Text = 'Local Admin Cleanup'
    $confirmation.StartPosition = 'CenterParent'
    $confirmation.Size = New-Object System.Drawing.Size(980, 600)
    $confirmation.MinimumSize = New-Object System.Drawing.Size(820, 500)
    $confirmation.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $confirmation.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $confirmation.ShowInTaskbar = $false

    $heading = New-Object System.Windows.Forms.Label
    $heading.Text = 'Confirm local Administrators group cleanup'
    $heading.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $heading.ForeColor = [System.Drawing.Color]::FromArgb(126, 34, 206)
    $heading.Location = New-Object System.Drawing.Point(20, 14)
    $heading.AutoSize = $true
    $confirmation.Controls.Add($heading)

    $details = New-Object System.Windows.Forms.Label
    $details.Text = "Instance token: $Instance`r`nOnly the group principals listed below will be removed. Membership is revalidated before and after every removal."
    $details.Location = New-Object System.Drawing.Point(22, 48)
    $details.Size = New-Object System.Drawing.Size(920, 38)
    $details.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    $confirmation.Controls.Add($details)

    $targetList = New-Object System.Windows.Forms.ListView
    $targetList.Location = New-Object System.Drawing.Point(20, 92)
    $targetList.Size = New-Object System.Drawing.Size(922, 410)
    $targetList.View = [System.Windows.Forms.View]::Details
    $targetList.FullRowSelect = $true
    $targetList.GridLines = $true
    $targetList.HideSelection = $false
    $targetList.ShowItemToolTips = $true
    $targetList.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    [void]$targetList.Columns.Add('Server', 155)
    [void]$targetList.Columns.Add('Administrators group', 175)
    [void]$targetList.Columns.Add('Group principal', 220)
    [void]$targetList.Columns.Add('Exact ADSI path', 350)

    foreach ($target in @($LocalAdminTargets | Sort-Object Server, Name, Path)) {
        $item = New-Object System.Windows.Forms.ListViewItem([string]$target.Server)
        [void]$item.SubItems.Add([string]$target.Group)
        [void]$item.SubItems.Add([string]$target.Name)
        [void]$item.SubItems.Add([string]$target.Path)
        $item.ToolTipText = [string]$target.Path
        [void]$targetList.Items.Add($item)
    }
    $confirmation.Controls.Add($targetList)

    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Text = '{0} group principal(s) will be removed.' -f @($LocalAdminTargets).Count
    $countLabel.Location = New-Object System.Drawing.Point(22, 512)
    $countLabel.Size = New-Object System.Drawing.Size(500, 30)
    $countLabel.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left
    )
    $countLabel.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    $confirmation.Controls.Add($countLabel)

    $cancelLocalAdminButton = New-Object System.Windows.Forms.Button
    $cancelLocalAdminButton.Text = 'Cancel'
    $cancelLocalAdminButton.Location = New-Object System.Drawing.Point(670, 520)
    $cancelLocalAdminButton.Size = New-Object System.Drawing.Size(105, 32)
    $cancelLocalAdminButton.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $cancelLocalAdminButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $confirmation.CancelButton = $cancelLocalAdminButton
    $confirmation.Controls.Add($cancelLocalAdminButton)

    $removeLocalAdminButton = New-Object System.Windows.Forms.Button
    $removeLocalAdminButton.Text = 'Remove Listed Groups'
    $removeLocalAdminButton.Location = New-Object System.Drawing.Point(782, 520)
    $removeLocalAdminButton.Size = New-Object System.Drawing.Size(160, 32)
    $removeLocalAdminButton.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $removeLocalAdminButton.BackColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
    $removeLocalAdminButton.ForeColor = [System.Drawing.Color]::White
    $removeLocalAdminButton.FlatStyle = 'Flat'
    $removeLocalAdminButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $confirmation.Controls.Add($removeLocalAdminButton)

    if (-not [string]::IsNullOrWhiteSpace($RenderPreviewPath)) {
        $confirmation.Show()
        [System.Windows.Forms.Application]::DoEvents()
        $bitmap = New-Object System.Drawing.Bitmap($confirmation.Width, $confirmation.Height)
        try {
            $confirmation.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, $confirmation.Width, $confirmation.Height)))
            $bitmap.Save($RenderPreviewPath, [System.Drawing.Imaging.ImageFormat]::Png)
        } finally {
            $bitmap.Dispose()
            $confirmation.Close()
            $confirmation.Dispose()
        }
        return $false
    }

    try {
        return ($confirmation.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::Yes)
    } finally {
        $confirmation.Dispose()
    }
}

function Show-ODBCCleanupConfirmation {
    param(
        [object[]]$ODBCTargets,
        [string]$Instance,
        [string]$RenderPreviewPath
    )

    $confirmation = New-Object System.Windows.Forms.Form
    $confirmation.Text = 'ODBC System DSN Cleanup'
    $confirmation.StartPosition = 'CenterParent'
    $confirmation.Size = New-Object System.Drawing.Size(1120, 600)
    $confirmation.MinimumSize = New-Object System.Drawing.Size(900, 500)
    $confirmation.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $confirmation.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $confirmation.ShowInTaskbar = $false

    $heading = New-Object System.Windows.Forms.Label
    $heading.Text = 'Confirm 32-bit and 64-bit System DSNs to delete'
    $heading.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $heading.ForeColor = [System.Drawing.Color]::FromArgb(153, 27, 27)
    $heading.Location = New-Object System.Drawing.Point(20, 14)
    $heading.AutoSize = $true
    $confirmation.Controls.Add($heading)

    $details = New-Object System.Windows.Forms.Label
    $details.Text = "Instance token: $Instance`r`nOnly the exact System DSNs listed below are eligible. Registry state and the token match are revalidated immediately before deletion."
    $details.Location = New-Object System.Drawing.Point(22, 48)
    $details.Size = New-Object System.Drawing.Size(1060, 38)
    $details.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    $confirmation.Controls.Add($details)

    $targetList = New-Object System.Windows.Forms.ListView
    $targetList.Location = New-Object System.Drawing.Point(20, 92)
    $targetList.Size = New-Object System.Drawing.Size(1062, 410)
    $targetList.View = [System.Windows.Forms.View]::Details
    $targetList.FullRowSelect = $true
    $targetList.GridLines = $true
    $targetList.HideSelection = $false
    $targetList.ShowItemToolTips = $true
    $targetList.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    [void]$targetList.Columns.Add('Server', 145)
    [void]$targetList.Columns.Add('Architecture', 95)
    [void]$targetList.Columns.Add('System DSN', 210)
    [void]$targetList.Columns.Add('Driver', 225)
    [void]$targetList.Columns.Add('Data server', 155)
    [void]$targetList.Columns.Add('Database', 140)
    [void]$targetList.Columns.Add('Registry path', 360)

    foreach ($target in @($ODBCTargets | Sort-Object Server, Architecture, Name)) {
        $item = New-Object System.Windows.Forms.ListViewItem([string]$target.Server)
        [void]$item.SubItems.Add([string]$target.Architecture)
        [void]$item.SubItems.Add([string]$target.Name)
        [void]$item.SubItems.Add([string]$target.Driver)
        [void]$item.SubItems.Add([string]$target.DataServer)
        [void]$item.SubItems.Add([string]$target.Database)
        [void]$item.SubItems.Add([string]$target.RegistryPath)
        $item.ToolTipText = '{0} | {1} | {2}' -f $target.Server, $target.Architecture, $target.RegistryPath
        [void]$targetList.Items.Add($item)
    }
    $confirmation.Controls.Add($targetList)

    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Text = '{0} System DSN target(s) will be deleted.' -f @($ODBCTargets).Count
    $countLabel.Location = New-Object System.Drawing.Point(22, 512)
    $countLabel.Size = New-Object System.Drawing.Size(500, 30)
    $countLabel.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left)
    $countLabel.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    $confirmation.Controls.Add($countLabel)

    $cancelODBCButton = New-Object System.Windows.Forms.Button
    $cancelODBCButton.Text = 'Cancel'
    $cancelODBCButton.Location = New-Object System.Drawing.Point(827, 520)
    $cancelODBCButton.Size = New-Object System.Drawing.Size(105, 32)
    $cancelODBCButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    $cancelODBCButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $confirmation.CancelButton = $cancelODBCButton
    $confirmation.Controls.Add($cancelODBCButton)

    $deleteODBCButton = New-Object System.Windows.Forms.Button
    $deleteODBCButton.Text = 'Delete Listed DSNs'
    $deleteODBCButton.Location = New-Object System.Drawing.Point(939, 520)
    $deleteODBCButton.Size = New-Object System.Drawing.Size(143, 32)
    $deleteODBCButton.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right)
    $deleteODBCButton.BackColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
    $deleteODBCButton.ForeColor = [System.Drawing.Color]::White
    $deleteODBCButton.FlatStyle = 'Flat'
    $deleteODBCButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $confirmation.Controls.Add($deleteODBCButton)

    if (-not [string]::IsNullOrWhiteSpace($RenderPreviewPath)) {
        $confirmation.Show()
        [System.Windows.Forms.Application]::DoEvents()
        $bitmap = New-Object System.Drawing.Bitmap($confirmation.Width, $confirmation.Height)
        try {
            $confirmation.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, $confirmation.Width, $confirmation.Height)))
            $bitmap.Save($RenderPreviewPath, [System.Drawing.Imaging.ImageFormat]::Png)
        } finally {
            $bitmap.Dispose()
            $confirmation.Close()
            $confirmation.Dispose()
        }
        return $false
    }

    try {
        return ($confirmation.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::Yes)
    } finally {
        $confirmation.Dispose()
    }
}

function Show-IISDeletionConfirmation {
    param(
        [object[]]$IISTargets,
        [string]$Instance,
        [string]$MatchMode,
        [string]$RenderPreviewPath
    )

    $plannedTargets = @($IISTargets | Where-Object { [bool]$_.Planned })
    $blockedTargets = @($IISTargets | Where-Object { -not [bool]$_.Planned })

    $confirmation = New-Object System.Windows.Forms.Form
    $confirmation.Text = 'IIS Pool and Site Decom'
    $confirmation.StartPosition = 'CenterParent'
    $confirmation.Size = New-Object System.Drawing.Size(1180, 650)
    $confirmation.MinimumSize = New-Object System.Drawing.Size(940, 520)
    $confirmation.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $confirmation.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $confirmation.ShowInTaskbar = $false

    $heading = New-Object System.Windows.Forms.Label
    $heading.Text = 'Confirm IIS targets in the removal plan'
    $heading.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $heading.ForeColor = [System.Drawing.Color]::FromArgb(153, 27, 27)
    $heading.Location = New-Object System.Drawing.Point(20, 14)
    $heading.AutoSize = $true
    $confirmation.Controls.Add($heading)

    $details = New-Object System.Windows.Forms.Label
    $details.Text = "Instance: $Instance | Match mode: $MatchMode`r`nDELETE rows are confirmed targets. BLOCKED pools remain untouched. Default Web Site and physical files are preserved."
    $details.Location = New-Object System.Drawing.Point(22, 48)
    $details.Size = New-Object System.Drawing.Size(1120, 38)
    $details.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    $confirmation.Controls.Add($details)

    $iisList = New-Object System.Windows.Forms.ListView
    $iisList.Location = New-Object System.Drawing.Point(20, 92)
    $iisList.Size = New-Object System.Drawing.Size(1122, 455)
    $iisList.View = [System.Windows.Forms.View]::Details
    $iisList.FullRowSelect = $true
    $iisList.GridLines = $true
    $iisList.HideSelection = $false
    $iisList.ShowItemToolTips = $true
    $iisList.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    [void]$iisList.Columns.Add('Server', 125)
    [void]$iisList.Columns.Add('Plan', 75)
    [void]$iisList.Columns.Add('Type', 90)
    [void]$iisList.Columns.Add('Name', 185)
    [void]$iisList.Columns.Add('Parent site', 160)
    [void]$iisList.Columns.Add('Application path', 170)
    [void]$iisList.Columns.Add('Pool', 170)
    [void]$iisList.Columns.Add('State', 75)
    [void]$iisList.Columns.Add('Details / blocked reason', 330)

    foreach ($target in @($IISTargets | Sort-Object Server, TargetType, Name)) {
        $planText = if ([bool]$target.Planned) { 'DELETE' } else { 'BLOCKED' }
        $detailText = if ([bool]$target.Planned) {
            [string]$target.Details
        } else {
            [string]$target.BlockedReason
        }

        $item = New-Object System.Windows.Forms.ListViewItem([string]$target.Server)
        [void]$item.SubItems.Add($planText)
        [void]$item.SubItems.Add([string]$target.TargetType)
        [void]$item.SubItems.Add([string]$target.Name)
        [void]$item.SubItems.Add([string]$target.SiteName)
        [void]$item.SubItems.Add([string]$target.ApplicationPath)
        [void]$item.SubItems.Add([string]$target.AppPoolName)
        [void]$item.SubItems.Add([string]$target.State)
        [void]$item.SubItems.Add($detailText)
        $item.ToolTipText = $detailText

        if ([bool]$target.Planned) {
            $item.BackColor = [System.Drawing.Color]::FromArgb(240, 253, 244)
            $item.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
        } else {
            $item.BackColor = [System.Drawing.Color]::FromArgb(254, 242, 242)
            $item.ForeColor = [System.Drawing.Color]::FromArgb(153, 27, 27)
        }
        [void]$iisList.Items.Add($item)
    }
    $confirmation.Controls.Add($iisList)

    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Text = '{0} planned target(s); {1} blocked pool(s).' -f $plannedTargets.Count, $blockedTargets.Count
    $countLabel.Location = New-Object System.Drawing.Point(22, 558)
    $countLabel.Size = New-Object System.Drawing.Size(520, 30)
    $countLabel.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left
    )
    $countLabel.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    $confirmation.Controls.Add($countLabel)

    $cancelIISButton = New-Object System.Windows.Forms.Button
    $cancelIISButton.Text = 'Cancel'
    $cancelIISButton.Location = New-Object System.Drawing.Point(832, 570)
    $cancelIISButton.Size = New-Object System.Drawing.Size(105, 32)
    $cancelIISButton.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $cancelIISButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $confirmation.CancelButton = $cancelIISButton
    $confirmation.Controls.Add($cancelIISButton)

    $deleteIISButton = New-Object System.Windows.Forms.Button
    $deleteIISButton.Text = 'Delete Planned IIS Targets'
    $deleteIISButton.Location = New-Object System.Drawing.Point(944, 570)
    $deleteIISButton.Size = New-Object System.Drawing.Size(198, 32)
    $deleteIISButton.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $deleteIISButton.BackColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
    $deleteIISButton.ForeColor = [System.Drawing.Color]::White
    $deleteIISButton.FlatStyle = 'Flat'
    $deleteIISButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $deleteIISButton.Enabled = $plannedTargets.Count -gt 0
    $confirmation.Controls.Add($deleteIISButton)

    if (-not [string]::IsNullOrWhiteSpace($RenderPreviewPath)) {
        $confirmation.Show()
        [System.Windows.Forms.Application]::DoEvents()
        $bitmap = New-Object System.Drawing.Bitmap($confirmation.Width, $confirmation.Height)
        try {
            $confirmation.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, $confirmation.Width, $confirmation.Height)))
            $bitmap.Save($RenderPreviewPath, [System.Drawing.Imaging.ImageFormat]::Png)
        } finally {
            $bitmap.Dispose()
            $confirmation.Close()
            $confirmation.Dispose()
        }
        return $false
    }

    try {
        return ($confirmation.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::Yes)
    } finally {
        $confirmation.Dispose()
    }
}

function Show-FolderSelectionDialog {
    param(
        [object[]]$FolderTargets,
        [bool]$HadDiscoveryErrors = $false,
        [string]$InstanceToken,
        [string]$RenderPreviewPath
    )

    $picker = New-Object System.Windows.Forms.Form
    $picker.Text = 'Choose Client Folders'
    $picker.StartPosition = 'CenterParent'
    $picker.Size = New-Object System.Drawing.Size(860, 600)
    $picker.MinimumSize = New-Object System.Drawing.Size(700, 480)
    $picker.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $picker.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $picker.ShowInTaskbar = $false

    $heading = New-Object System.Windows.Forms.Label
    $heading.Text = 'Select the exact server/folder targets'
    $heading.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $heading.Location = New-Object System.Drawing.Point(18, 14)
    $heading.AutoSize = $true
    $picker.Controls.Add($heading)

    $instructions = New-Object System.Windows.Forms.Label
    $instructions.Text = 'Only checked rows will be eligible for deletion. Nothing is deleted from this window.'
    $instructions.Location = New-Object System.Drawing.Point(20, 48)
    $instructions.AutoSize = $true
    $instructions.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    $picker.Controls.Add($instructions)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(20, 78)
    $listView.Size = New-Object System.Drawing.Size(802, 405)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.CheckBoxes = $true
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.HideSelection = $false
    $listView.ShowItemToolTips = $true
    $listView.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    [void]$listView.Columns.Add('Server', 180)
    [void]$listView.Columns.Add('Full UNC path', 590)

    foreach ($target in @($FolderTargets | Sort-Object Server, Name)) {
        $item = New-Object System.Windows.Forms.ListViewItem([string]$target.Server)
        $uncPath = '\\{0}\m$\OBOL\Clients\{1}' -f $target.Server, $target.Name
        $pathSubItem = $item.SubItems.Add($uncPath)
        if (
            -not [string]::IsNullOrWhiteSpace($InstanceToken) -and
            $uncPath.IndexOf(
                $InstanceToken.Trim(),
                [System.StringComparison]::OrdinalIgnoreCase
            ) -ge 0
        ) {
            $item.UseItemStyleForSubItems = $false
            $item.BackColor = [System.Drawing.Color]::FromArgb(220, 252, 231)
            $item.ForeColor = [System.Drawing.Color]::FromArgb(21, 128, 61)
            $pathSubItem.BackColor = [System.Drawing.Color]::FromArgb(220, 252, 231)
            $pathSubItem.ForeColor = [System.Drawing.Color]::FromArgb(21, 128, 61)
        } else {
            $item.UseItemStyleForSubItems = $false
            $item.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
            $item.ForeColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
            $pathSubItem.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
            $pathSubItem.ForeColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
        }
        $item.ToolTipText = $uncPath
        $item.Tag = $target
        [void]$listView.Items.Add($item)
    }
    $picker.Controls.Add($listView)

    $status = New-Object System.Windows.Forms.Label
    $status.Location = New-Object System.Drawing.Point(20, 492)
    $status.Size = New-Object System.Drawing.Size(570, 35)
    $status.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left
    )
    if ($HadDiscoveryErrors) {
        $status.Text = 'Some servers could not be listed. Review the background activity log.'
        $status.ForeColor = [System.Drawing.Color]::FromArgb(180, 83, 9)
    } else {
        $status.Text = '{0} folder target(s) available.' -f @($FolderTargets).Count
        $status.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
    }
    if (-not [string]::IsNullOrWhiteSpace($InstanceToken)) {
        $status.Text += " Green paths contain instance token '$($InstanceToken.Trim())'; red paths do not."
    } else {
        $status.Text += ' Red paths do not match an instance token.'
    }
    $picker.Controls.Add($status)

    $checkAllButton = New-Object System.Windows.Forms.Button
    $checkAllButton.Text = 'Check all'
    $checkAllButton.Location = New-Object System.Drawing.Point(20, 530)
    $checkAllButton.Size = New-Object System.Drawing.Size(95, 30)
    $checkAllButton.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left
    )
    $checkAllButton.Add_Click({
        foreach ($item in $listView.Items) {
            $item.Checked = $true
        }
    })
    $picker.Controls.Add($checkAllButton)

    $clearAllButton = New-Object System.Windows.Forms.Button
    $clearAllButton.Text = 'Clear all'
    $clearAllButton.Location = New-Object System.Drawing.Point(122, 530)
    $clearAllButton.Size = New-Object System.Drawing.Size(95, 30)
    $clearAllButton.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left
    )
    $clearAllButton.Add_Click({
        foreach ($item in $listView.Items) {
            $item.Checked = $false
        }
    })
    $picker.Controls.Add($clearAllButton)

    $openPathsButton = New-Object System.Windows.Forms.Button
    $openPathsButton.Text = 'Open Selected Paths'
    $openPathsButton.Location = New-Object System.Drawing.Point(224, 530)
    $openPathsButton.Size = New-Object System.Drawing.Size(165, 30)
    $openPathsButton.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left
    )
    $openPathsButton.Add_Click({
        $itemsToOpen = if ($listView.CheckedItems.Count -gt 0) {
            @($listView.CheckedItems)
        } else {
            @($listView.SelectedItems)
        }

        if ($itemsToOpen.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                $picker,
                'Check or highlight at least one UNC path.',
                'No paths selected',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        if ($itemsToOpen.Count -gt 5) {
            $openChoice = [System.Windows.Forms.MessageBox]::Show(
                $picker,
                ("Open {0} separate File Explorer windows?" -f $itemsToOpen.Count),
                'Open UNC paths',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button2
            )
            if ($openChoice -ne [System.Windows.Forms.DialogResult]::Yes) {
                return
            }
        }

        $openFailures = New-Object System.Collections.ArrayList
        foreach ($itemToOpen in $itemsToOpen) {
            $uncPathToOpen = [string]$itemToOpen.SubItems[1].Text
            try {
                Start-Process `
                    -FilePath 'explorer.exe' `
                    -ArgumentList ('"{0}"' -f $uncPathToOpen) `
                    -ErrorAction Stop
            } catch {
                [void]$openFailures.Add(
                    ("{0}: {1}" -f $uncPathToOpen, $_.Exception.Message)
                )
            }
        }

        if ($openFailures.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show(
                $picker,
                ("Some paths could not be opened:`r`n`r`n{0}" -f ($openFailures -join "`r`n")),
                'File Explorer',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }
    })
    $picker.Controls.Add($openPathsButton)

    $useSelectionButton = New-Object System.Windows.Forms.Button
    $useSelectionButton.Text = 'Use Selected Folders'
    $useSelectionButton.Location = New-Object System.Drawing.Point(642, 530)
    $useSelectionButton.Size = New-Object System.Drawing.Size(180, 30)
    $useSelectionButton.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $useSelectionButton.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
    $useSelectionButton.ForeColor = [System.Drawing.Color]::White
    $useSelectionButton.FlatStyle = 'Flat'
    $useSelectionButton.Add_Click({
        if ($listView.CheckedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                $picker,
                'Check at least one server/folder row.',
                'No folders selected',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $picker.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $picker.Close()
    })
    $picker.Controls.Add($useSelectionButton)

    $cancelSelectionButton = New-Object System.Windows.Forms.Button
    $cancelSelectionButton.Text = 'Cancel'
    $cancelSelectionButton.Location = New-Object System.Drawing.Point(535, 530)
    $cancelSelectionButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelSelectionButton.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $cancelSelectionButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $picker.CancelButton = $cancelSelectionButton
    $picker.Controls.Add($cancelSelectionButton)

    if (-not [string]::IsNullOrWhiteSpace($RenderPreviewPath)) {
        $picker.Show()
        [System.Windows.Forms.Application]::DoEvents()
        $bitmap = New-Object System.Drawing.Bitmap($picker.Width, $picker.Height)
        try {
            $picker.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, $picker.Width, $picker.Height)))
            $bitmap.Save($RenderPreviewPath, [System.Drawing.Imaging.ImageFormat]::Png)
        } finally {
            $bitmap.Dispose()
            $picker.Close()
            $picker.Dispose()
        }
        return @()
    }

    try {
        $result = $picker.ShowDialog($form)
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            return @()
        }

        return @(
            $listView.CheckedItems |
                ForEach-Object { $_.Tag }
        )
    } finally {
        $picker.Dispose()
    }
}

function Confirm-FolderDeletion {
    param([object[]]$FolderTargets)

    $confirmation = New-Object System.Windows.Forms.Form
    $confirmation.Text = 'Confirm Folder Deletion'
    $confirmation.StartPosition = 'CenterParent'
    $confirmation.Size = New-Object System.Drawing.Size(740, 520)
    $confirmation.MinimumSize = New-Object System.Drawing.Size(620, 440)
    $confirmation.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $confirmation.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $confirmation.ShowInTaskbar = $false

    $warning = New-Object System.Windows.Forms.Label
    $warning.Text = 'The following folders will be recursively and permanently deleted:'
    $warning.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
    $warning.ForeColor = [System.Drawing.Color]::FromArgb(153, 27, 27)
    $warning.Location = New-Object System.Drawing.Point(18, 16)
    $warning.AutoSize = $true
    $confirmation.Controls.Add($warning)

    $targetTextBox = New-Object System.Windows.Forms.TextBox
    $targetTextBox.Location = New-Object System.Drawing.Point(20, 50)
    $targetTextBox.Size = New-Object System.Drawing.Size(682, 365)
    $targetTextBox.Multiline = $true
    $targetTextBox.ReadOnly = $true
    $targetTextBox.ScrollBars = 'Both'
    $targetTextBox.WordWrap = $false
    $targetTextBox.Font = New-Object System.Drawing.Font('Consolas', 9)
    $targetTextBox.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $targetTextBox.Text = @(
        $FolderTargets |
            Sort-Object Server, Name |
            ForEach-Object {
                '\\{0}\m$\OBOL\Clients\{1}' -f $_.Server, $_.Name
            }
    ) -join [Environment]::NewLine
    $confirmation.Controls.Add($targetTextBox)

    $cancelButtonLocal = New-Object System.Windows.Forms.Button
    $cancelButtonLocal.Text = 'Cancel'
    $cancelButtonLocal.Location = New-Object System.Drawing.Point(400, 435)
    $cancelButtonLocal.Size = New-Object System.Drawing.Size(105, 32)
    $cancelButtonLocal.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $cancelButtonLocal.DialogResult = [System.Windows.Forms.DialogResult]::No
    $confirmation.CancelButton = $cancelButtonLocal
    $confirmation.Controls.Add($cancelButtonLocal)

    $deleteButtonLocal = New-Object System.Windows.Forms.Button
    $deleteButtonLocal.Text = 'Delete Selected Folders'
    $deleteButtonLocal.Location = New-Object System.Drawing.Point(512, 435)
    $deleteButtonLocal.Size = New-Object System.Drawing.Size(190, 32)
    $deleteButtonLocal.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $deleteButtonLocal.BackColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
    $deleteButtonLocal.ForeColor = [System.Drawing.Color]::White
    $deleteButtonLocal.FlatStyle = 'Flat'
    $deleteButtonLocal.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $confirmation.Controls.Add($deleteButtonLocal)

    try {
        return ($confirmation.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::Yes)
    } finally {
        $confirmation.Dispose()
    }
}

function ConvertTo-OperationReportTarget {
    param([Parameter(Mandatory)]$Record)

    $category = [string]$Record.Category
    $action = switch ($category) {
        'Services' {
            if ([string]$Record.TargetType -in @('GCS list entry', 'GCSListEntry')) { 'Remove matching GCS list entry' } else { 'Stop and delete Windows service' }
        }
        'IISSites' { 'Stop and remove IIS site' }
        'ApplicationPools' { 'Stop and remove IIS application pool after dependency recheck' }
        'ODBCDataSources' { 'Remove exact System DSN from selected registry view' }
        'LocalAdmins' { 'Remove principal from local Administrators group' }
        'Folders' { 'Delete exact client folder' }
        default { 'Remove selected target' }
    }
    $previousState = if ($category -eq 'ODBCDataSources') { 'Present' } else { [string]$Record.State }
    if ([string]::IsNullOrWhiteSpace($previousState)) {
        $previousState = switch ($category) {
            'Folders' { 'Present' }
            'LocalAdmins' { 'Member' }
            'ODBCDataSources' { 'Present' }
            default { 'Discovered' }
        }
    }

    [PSCustomObject]@{
        Category      = $category
        Server        = [string]$Record.Server
        TargetType    = [string]$Record.TargetType
        Name          = [string]$Record.Name
        Architecture  = [string]$Record.Architecture
        Path          = [string]$Record.Path
        PlannedAction = $action
        PreviousState = $previousState
    }
}

function ConvertTo-ActionReportTargets {
    param(
        [Parameter(Mandatory)][object[]]$Records,
        [Parameter(Mandatory)][string]$DefaultCategory
    )

    @(
        foreach ($record in @($Records)) {
            $category = if ($DefaultCategory -eq 'IIS') {
                switch ([string]$record.TargetType) {
                    'AppPool' { 'ApplicationPools' }
                    default { 'IISSites' }
                }
            } else {
                $DefaultCategory
            }
            $path = [string]$record.Path
            if ($category -eq 'Folders' -and [string]::IsNullOrWhiteSpace($path)) {
                $path = '\\{0}\m$\OBOL\Clients\{1}' -f $record.Server, $record.Name
            }
            ConvertTo-OperationReportTarget -Record ([PSCustomObject]@{
                Category = $category
                Server = [string]$record.Server
                TargetType = [string]$record.TargetType
                Name = [string]$record.Name
                Architecture = [string]$record.Architecture
                Path = $path
                State = [string]$record.State
            })
        }
    )
}

function Confirm-GuiAction {
    param(
        [string]$Message,
        [string]$Title
    )

    $choice = [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    )
    return ($choice -eq [System.Windows.Forms.DialogResult]::Yes)
}

function ConvertTo-OperationReportRows {
    param(
        [Parameter(Mandatory)][object[]]$Steps,
        [object[]]$StepResults,
        [Parameter(Mandatory)][string]$SessionId,
        [Parameter(Mandatory)][string]$ActionName,
        [Parameter(Mandatory)][datetime]$StartedAt,
        [Parameter(Mandatory)][datetime]$EndedAt,
        [bool]$Cancelled
    )

    $resultsByStep = @{}
    foreach ($result in @($StepResults)) {
        $resultsByStep[[string]$result.StepId] = $result
    }

    @(
        foreach ($step in @($Steps)) {
            $stepResult = $resultsByStep[[string]$step.StepId]
            $outcome = if ($Cancelled) {
                'Cancelled'
            } elseif ($null -eq $stepResult) {
                'Not completed'
            } elseif ([int]$stepResult.ExitCode -eq 0) {
                'Succeeded'
            } else {
                'Failed'
            }
            $targets = @($step.ReportTargets)
            if ($targets.Count -eq 0) {
                $targets = @([PSCustomObject]@{
                    Category = [string]$step.Category; Server = [string]$step.Server
                    TargetType = 'Worker step'; Name = [string]$step.Title
                    Architecture = ''; Path = ''; PlannedAction = [string]$step.Title
                    PreviousState = 'Discovered'
                })
            }

            foreach ($target in $targets) {
                [PSCustomObject]@{
                    SessionId       = $SessionId
                    Operator        = [Environment]::UserName
                    Action          = $ActionName
                    StartedAt       = $StartedAt.ToString('yyyy-MM-dd HH:mm:ss')
                    EndedAt         = $EndedAt.ToString('yyyy-MM-dd HH:mm:ss')
                    DurationSeconds = [math]::Round(($EndedAt - $StartedAt).TotalSeconds, 2)
                    Category        = [string]$target.Category
                    Server          = [string]$target.Server
                    TargetType      = [string]$target.TargetType
                    TargetName      = [string]$target.Name
                    Architecture    = [string]$target.Architecture
                    Path            = [string]$target.Path
                    PreviousState   = [string]$target.PreviousState
                    PlannedAction   = [string]$target.PlannedAction
                    Outcome         = $outcome
                    FinalState      = switch ($outcome) {
                        'Succeeded' { 'Removed or already absent (worker verified)' }
                        'Failed' { 'Removal incomplete; review message and rescan' }
                        'Cancelled' { 'Not changed by the cancelled action' }
                        default { 'Unknown; run a verification scan' }
                    }
                    ExitCode        = if ($null -eq $stepResult) { '' } else { [string]$stepResult.ExitCode }
                    ErrorMessage    = if ($null -ne $stepResult -and [int]$stepResult.ExitCode -ne 0) { [string]$stepResult.Message } else { '' }
                    Message         = if ($null -eq $stepResult) { 'The worker step did not return a completion record.' } else { [string]$stepResult.Message }
                }
            }
        }
    )
}

function Write-OperationReport {
    param(
        [object[]]$Steps,
        [object[]]$StepResults,
        [string]$SessionId,
        [string]$ActionName,
        [datetime]$StartedAt,
        [datetime]$EndedAt,
        [bool]$Cancelled
    )

    $reportRows = @(ConvertTo-OperationReportRows -Steps $Steps -StepResults $StepResults `
        -SessionId $SessionId -ActionName $ActionName -StartedAt $StartedAt -EndedAt $EndedAt -Cancelled $Cancelled)
    if ($reportRows.Count -eq 0) { return '' }

    $reportDirectory = Join-Path $PSScriptRoot 'Logs'
    if (-not (Test-Path -LiteralPath $reportDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
    }
    $safeActionName = $ActionName -replace '[^A-Za-z0-9._-]', '_'
    $reportPath = Join-Path $reportDirectory ('Operation_{0}_{1}_{2}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), $safeActionName, $SessionId)
    $reportRows | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8 -Force
    return $reportPath
}

function New-WorkerStep {
    param(
        [string]$Title,
        [string]$ScriptPath,
        [string[]]$Arguments,
        [string]$Category = '',
        [string]$Server = '',
        [object[]]$ReportTargets = @()
    )

    [PSCustomObject]@{
        StepId     = [guid]::NewGuid().ToString('N')
        Title      = $Title
        ScriptPath = $ScriptPath
        Arguments  = @($Arguments)
        Category   = $Category
        Server     = $Server
        ReportTargets = @($ReportTargets)
    }
}

$script:BackgroundRunner = {
    param($Steps)

    $powerShellExe = Join-Path -Path $PSHOME -ChildPath 'powershell.exe'
    foreach ($step in $Steps) {
        $stepStartedAt = Get-Date
        Write-Output ('@@STEP_START@@' + ([PSCustomObject]@{
            StepId = [string]$step.StepId; Title = [string]$step.Title
            Category = [string]$step.Category; Server = [string]$step.Server
            StartedAt = $stepStartedAt.ToString('o')
        } | ConvertTo-Json -Compress))
        Write-Output ("=== {0} ===" -f $step.Title)

        $arguments = @(
            '-NoLogo'
            '-NoProfile'
            '-NonInteractive'
            '-ExecutionPolicy'
            'Bypass'
            '-File'
            [string]$step.ScriptPath
        )
        $arguments += @($step.Arguments)

        try {
            & $powerShellExe @arguments 2>&1 |
                ForEach-Object { Write-Output ([string]$_) }
            $exitCode = $LASTEXITCODE
        } catch {
            Write-Output ("ERROR: Unable to launch worker: {0}" -f $_.Exception.Message)
            $exitCode = 1
        }

        if ($exitCode -eq 0) {
            Write-Output ("=== {0} finished successfully ===" -f $step.Title)
        } else {
            Write-Output ("ERROR: {0} finished with exit code {1}" -f $step.Title, $exitCode)
        }
        Write-Output ('@@STEP_RESULT@@' + ([PSCustomObject]@{
            StepId = [string]$step.StepId; Title = [string]$step.Title
            Category = [string]$step.Category; Server = [string]$step.Server
            StartedAt = $stepStartedAt.ToString('o'); EndedAt = (Get-Date).ToString('o')
            ExitCode = [int]$exitCode
            Message = if ($exitCode -eq 0) { 'Worker completed successfully.' } else { "Worker finished with exit code $exitCode." }
        } | ConvertTo-Json -Compress))
    }
}

$script:FolderDiscoveryRunner = {
    param([string[]]$Servers)

    $remoteListScript = {
        $basePath = 'M:\OBOL\Clients'
        if (-not (Test-Path -LiteralPath $basePath -PathType Container -ErrorAction Stop)) {
            throw "Client folder root does not exist: $basePath"
        }

        @(
            Get-ChildItem -LiteralPath $basePath -Directory -Force -ErrorAction Stop |
                ForEach-Object { $_.Name } |
                Sort-Object -Unique
        )
    }

    foreach ($server in $Servers) {
        Write-Output ([PSCustomObject]@{
            Kind    = 'Log'
            Message = "Listing client folders on $server..."
        })

        $names = @()
        $method = 'Remote'
        try {
            $names = @(
                Invoke-Command -ComputerName $server -ScriptBlock $remoteListScript -ErrorAction Stop
            )
        } catch {
            Write-Output ([PSCustomObject]@{
                Kind    = 'Warning'
                Message = "WARNING: Remote folder listing failed on $server`: $($_.Exception.Message). Trying UNC."
            })

            $method = 'UNC'
            $uncRoot = "\\$server\m$\OBOL\Clients"
            try {
                if (-not (Test-Path -LiteralPath $uncRoot -PathType Container -ErrorAction Stop)) {
                    throw "Client folder root does not exist: $uncRoot"
                }
                $names = @(
                    Get-ChildItem -LiteralPath $uncRoot -Directory -Force -ErrorAction Stop |
                        ForEach-Object { $_.Name } |
                        Sort-Object -Unique
                )
            } catch {
                Write-Output ([PSCustomObject]@{
                    Kind    = 'Error'
                    Message = "ERROR: Unable to list client folders on $server`: $($_.Exception.Message)"
                })
                continue
            }
        }

        foreach ($name in @($names | ForEach-Object { [string]$_ } | Where-Object { $_ } | Sort-Object -Unique)) {
            Write-Output ([PSCustomObject]@{
                Kind   = 'Folder'
                Server = [string]$server
                Name   = $name
            })
        }

        Write-Output ([PSCustomObject]@{
            Kind    = 'Log'
            Message = "Found $(@($names).Count) client folder(s) on $server using $method."
        })
    }
}

function New-StreamingInvocation {
    param(
        [Parameter(Mandatory)][System.Management.Automation.PowerShell]$PowerShell,
        [Parameter(Mandatory)]$Output
    )

    $beginMethod = @(
        [System.Management.Automation.PowerShell].GetMethods() |
            Where-Object {
                $_.Name -eq 'BeginInvoke' -and
                $_.IsGenericMethodDefinition -and
                $_.GetGenericArguments().Count -eq 2 -and
                $_.GetParameters().Count -eq 2
            }
    )[0]
    if ($null -eq $beginMethod) {
        throw 'The streaming BeginInvoke overload is unavailable.'
    }

    $closedBeginMethod = $beginMethod.MakeGenericMethod(
        [Type[]]@([object], [psobject])
    )
    return $closedBeginMethod.Invoke(
        $PowerShell,
        [object[]]@($null, $Output.PSObject.BaseObject)
    )
}

function Start-BackgroundSteps {
    param(
        [System.Collections.IList]$Steps,
        [string]$ActionName,
        [string]$JobKind = 'WorkerSteps',
        $Context = $null
    )

    if ($null -ne $script:CurrentJob) {
        Show-InputError 'Another action is already running.'
        return
    }
    if ($Steps.Count -eq 0) {
        Show-InputError 'No complete action inputs were provided.'
        return
    }

    $powerShell = [PowerShell]::Create()
    [void]$powerShell.AddScript($script:BackgroundRunner.ToString())
    [void]$powerShell.AddArgument($Steps)

    $output = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
    try {
        $asyncResult = New-StreamingInvocation -PowerShell $powerShell -Output $output
    } catch {
        $powerShell.Dispose()
        Add-ActivityLog -Message ("ERROR: Unable to start background action: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(248, 113, 113))
        return
    }

    $jobResults = $null
    if ($JobKind -in @('ServiceDiscovery', 'IISDiscovery', 'LocalAdminDiscovery', 'ODBCDiscovery', 'ScanDiscovery')) {
        $jobResults = New-Object System.Collections.ArrayList
    }

    $script:CurrentJob = [PSCustomObject]@{
        Name         = $ActionName
        Kind         = $JobKind
        PowerShell   = $powerShell
        Output       = $output
        OutputIndex  = 0
        ErrorIndex   = 0
        AsyncResult  = $asyncResult
        Cancelled    = $false
        Results      = $jobResults
        HadErrors    = $false
        Context      = $Context
        Steps        = @($Steps)
        StepResults  = New-Object System.Collections.ArrayList
        StartedAt    = Get-Date
        SessionId    = (Get-Date -Format 'yyyyMMddHHmmssfff')
        CompletedSteps = 0
        SeenServers  = @{}
    }

    Set-ActionState -Busy $true
    if ($JobKind -eq 'WorkerSteps') {
        $progressBar.Style = 'Continuous'
        $progressBar.Minimum = 0
        $progressBar.Maximum = [math]::Max(1, $Steps.Count)
        $progressBar.Value = 0
    }
    Set-StatusText -Text "Running: $ActionName"
    Add-ActivityLog -Message ("Started background action: {0}" -f $ActionName) -Color ([System.Drawing.Color]::FromArgb(103, 232, 249))
}

function Start-FolderDiscovery {
    param([string[]]$Servers)

    if ($null -ne $script:CurrentJob) {
        Show-InputError 'Another action is already running.'
        return
    }
    if (@($Servers).Count -eq 0) {
        Show-InputError 'Enter at least one service or application server.'
        return
    }

    $powerShell = [PowerShell]::Create()
    [void]$powerShell.AddScript($script:FolderDiscoveryRunner.ToString())
    [void]$powerShell.AddArgument($Servers)

    $output = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
    try {
        $asyncResult = New-StreamingInvocation -PowerShell $powerShell -Output $output
    } catch {
        $powerShell.Dispose()
        Add-ActivityLog -Message ("ERROR: Unable to start folder discovery: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(248, 113, 113))
        return
    }

    $script:CurrentJob = [PSCustomObject]@{
        Name         = 'Client folder discovery'
        Kind         = 'FolderDiscovery'
        PowerShell   = $powerShell
        Output       = $output
        OutputIndex  = 0
        ErrorIndex   = 0
        AsyncResult  = $asyncResult
        Cancelled    = $false
        Results      = New-Object System.Collections.ArrayList
        HadErrors    = $false
        Context      = [PSCustomObject]@{ ProgressServers = @($Servers) }
        SeenServers  = @{}
    }

    Set-ActionState -Busy $true
    Set-StatusText -Text 'Running: Client folder discovery'
    Add-ActivityLog -Message 'Started background action: Client folder discovery' -Color ([System.Drawing.Color]::FromArgb(103, 232, 249))
}

function Drain-CurrentJobOutput {
    if ($null -eq $script:CurrentJob) { return }

    while ($script:CurrentJob.OutputIndex -lt $script:CurrentJob.Output.Count) {
        $item = $script:CurrentJob.Output[$script:CurrentJob.OutputIndex]
        $script:CurrentJob.OutputIndex++

        $structuredMessage = [string]$item
        if ($structuredMessage.StartsWith('@@STEP_START@@', [System.StringComparison]::Ordinal)) {
            try {
                $stepStart = $structuredMessage.Substring('@@STEP_START@@'.Length) | ConvertFrom-Json -ErrorAction Stop
                $stepNumber = [math]::Min($script:CurrentJob.CompletedSteps + 1, @($script:CurrentJob.Steps).Count)
                $stepServers = @(ConvertTo-ServerArray -Text @([string]$stepStart.Server))
                $serverText = if ($stepServers.Count -eq 0) {
                    ''
                } elseif ($stepServers.Count -eq 1) {
                    " | Server: $($stepServers[0])"
                } else {
                    " | Servers: $($stepServers.Count)"
                }
                $fullServerText = if ([string]::IsNullOrWhiteSpace([string]$stepStart.Server)) {
                    ''
                } else {
                    " | Server(s): $($stepStart.Server)"
                }
                $shortStatus = 'Running step {0} of {1}: {2}{3}' -f $stepNumber, @($script:CurrentJob.Steps).Count, $stepStart.Title, $serverText
                $fullStatus = 'Running step {0} of {1}: {2}{3}' -f $stepNumber, @($script:CurrentJob.Steps).Count, $stepStart.Title, $fullServerText
                Set-StatusText -Text $shortStatus -FullText $fullStatus
            } catch {
                Add-ActivityLog -Message ("WARNING: Invalid step-start progress record: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(251, 191, 36))
            }
            continue
        }
        if ($structuredMessage.StartsWith('@@STEP_RESULT@@', [System.StringComparison]::Ordinal)) {
            try {
                $stepResult = $structuredMessage.Substring('@@STEP_RESULT@@'.Length) | ConvertFrom-Json -ErrorAction Stop
                [void]$script:CurrentJob.StepResults.Add($stepResult)
                $script:CurrentJob.CompletedSteps++
                if ($progressBar.Style -ne [System.Windows.Forms.ProgressBarStyle]::Continuous) {
                    $progressBar.Style = 'Continuous'
                }
                $progressBar.Maximum = [math]::Max(1, @($script:CurrentJob.Steps).Count)
                $progressBar.Value = [math]::Min($progressBar.Maximum, $script:CurrentJob.CompletedSteps)
            } catch {
                Add-ActivityLog -Message ("WARNING: Invalid step-result progress record: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(251, 191, 36))
            }
            continue
        }

        if ($structuredMessage -match '^====\s+([^=]+?)\s+::') {
            $currentServer = $matches[1].Trim()
            $progressServers = @($script:CurrentJob.Context.ProgressServers)
            if ($progressServers.Count -gt 0) {
                $serverKey = $currentServer.ToUpperInvariant()
                if (-not $script:CurrentJob.SeenServers.ContainsKey($serverKey)) {
                    $script:CurrentJob.SeenServers[$serverKey] = $true
                }
                $serverNumber = @($script:CurrentJob.SeenServers.Keys).Count
                Set-StatusText -Text ('Running: {0} | Server {1} of {2}: {3}' -f $script:CurrentJob.Name, $serverNumber, $progressServers.Count, $currentServer)
                $progressBar.Style = 'Continuous'
                $progressBar.Maximum = [math]::Max(1, $progressServers.Count)
                $progressBar.Value = [math]::Min($progressBar.Maximum, $serverNumber)
            } else {
                Set-StatusText -Text ('Running: {0} | Server: {1}' -f $script:CurrentJob.Name, $currentServer)
            }
        }

        if ($script:CurrentJob.Kind -eq 'ScanDiscovery') {
            $message = [string]$item
            $scanMarker = '@@SCAN_TARGET@@'
            if ($message.StartsWith($scanMarker, [System.StringComparison]::Ordinal)) {
                try {
                    $scanRecord = $message.Substring($scanMarker.Length) |
                        ConvertFrom-Json -ErrorAction Stop
                    $planned = [bool]$scanRecord.Planned
                    $blockedReason = [string]$scanRecord.BlockedReason
                    $details = [string]$scanRecord.Details
                    [void]$script:CurrentJob.Results.Add([PSCustomObject]@{
                        Category      = [string]$scanRecord.Category
                        Server        = [string]$scanRecord.Server
                        TargetType    = [string]$scanRecord.TargetType
                        Name          = [string]$scanRecord.Name
                        DisplayName   = [string]$scanRecord.DisplayName
                        State         = [string]$scanRecord.State
                        StartName     = [string]$scanRecord.StartName
                        FileName      = [string]$scanRecord.FileName
                        Path          = [string]$scanRecord.Path
                        Entry         = [string]$scanRecord.Entry
                        SiteName      = [string]$scanRecord.SiteName
                        AppPoolName   = [string]$scanRecord.AppPoolName
                        Architecture  = [string]$scanRecord.Architecture
                        Driver        = [string]$scanRecord.Driver
                        DataServer    = [string]$scanRecord.DataServer
                        Database      = [string]$scanRecord.Database
                        RegistryPath  = [string]$scanRecord.RegistryPath
                        Details       = $details
                        BlockedReason = $blockedReason
                        Planned       = $planned
                        MatchesToken  = [bool]$scanRecord.MatchesToken
                        MatchSource   = [string]$scanRecord.MatchSource
                        MatchReason   = [string]$scanRecord.MatchReason
                        PoolIdentity  = [string]$scanRecord.PoolIdentity
                        IdentityMatchesToken = [bool]$scanRecord.IdentityMatchesToken
                        IsOptInCandidate = [bool]$scanRecord.IsOptInCandidate
                        OptInSelected = $false
                        MatchText     = if ([bool]$scanRecord.MatchesToken) { 'Yes' } else { 'No' }
                        RemovalStatus = if ($planned) { 'Eligible' } else { 'Blocked' }
                        ScanDetails   = if (-not [string]::IsNullOrWhiteSpace($blockedReason)) {
                            $blockedReason
                        } elseif ([bool]$scanRecord.IsOptInCandidate) {
                            [string]$scanRecord.MatchReason
                        } else {
                            $details
                        }
                    })
                } catch {
                    $script:CurrentJob.HadErrors = $true
                    Add-ActivityLog -Message ("ERROR: Invalid scan result record: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(248, 113, 113))
                }
                continue
            }

            if ($message -match '(?i)^ERROR:|finished with exit code [1-9]') {
                $script:CurrentJob.HadErrors = $true
            }
            Add-ActivityLog -Message $message -Color (Get-LogColor -Message $message)
            continue
        }

        if ($script:CurrentJob.Kind -eq 'IISDiscovery') {
            $message = [string]$item
            $iisMarker = '@@IIS_TARGET@@'
            if ($message.StartsWith($iisMarker, [System.StringComparison]::Ordinal)) {
                try {
                    $iisRecord = $message.Substring($iisMarker.Length) |
                        ConvertFrom-Json -ErrorAction Stop
                    [void]$script:CurrentJob.Results.Add([PSCustomObject]@{
                        Server          = [string]$iisRecord.Server
                        TargetType      = [string]$iisRecord.TargetType
                        Name            = [string]$iisRecord.Name
                        SiteName        = [string]$iisRecord.SiteName
                        ApplicationPath = [string]$iisRecord.ApplicationPath
                        AppPoolName     = [string]$iisRecord.AppPoolName
                        State           = [string]$iisRecord.State
                        Planned         = [bool]$iisRecord.Planned
                        Details         = [string]$iisRecord.Details
                        BlockedReason   = [string]$iisRecord.BlockedReason
                        PoolIdentity    = [string]$iisRecord.PoolIdentity
                        IdentityMatchesToken = [bool]$iisRecord.IdentityMatchesToken
                        IsOptInCandidate = [bool]$iisRecord.IsOptInCandidate
                        MatchSource     = [string]$iisRecord.MatchSource
                        MatchReason     = [string]$iisRecord.MatchReason
                    })
                } catch {
                    $script:CurrentJob.HadErrors = $true
                    Add-ActivityLog -Message ("ERROR: Invalid IIS discovery record: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(248, 113, 113))
                }
                continue
            }

            if ($message -match '(?i)^ERROR:|finished with exit code [1-9]') {
                $script:CurrentJob.HadErrors = $true
            }
            Add-ActivityLog -Message $message -Color (Get-LogColor -Message $message)
            continue
        }

        if ($script:CurrentJob.Kind -eq 'ODBCDiscovery') {
            $message = [string]$item
            $odbcMarker = '@@ODBC_TARGET@@'
            if ($message.StartsWith($odbcMarker, [System.StringComparison]::Ordinal)) {
                try {
                    $odbcRecord = $message.Substring($odbcMarker.Length) |
                        ConvertFrom-Json -ErrorAction Stop
                    [void]$script:CurrentJob.Results.Add([PSCustomObject]@{
                        Server       = [string]$odbcRecord.Server
                        Architecture = [string]$odbcRecord.Architecture
                        Name         = [string]$odbcRecord.Name
                        Driver       = [string]$odbcRecord.Driver
                        KeyDriver    = [string]$odbcRecord.KeyDriver
                        DataServer   = [string]$odbcRecord.DataServer
                        Database     = [string]$odbcRecord.Database
                        Description  = [string]$odbcRecord.Description
                        RegistryPath = [string]$odbcRecord.RegistryPath
                    })
                } catch {
                    $script:CurrentJob.HadErrors = $true
                    Add-ActivityLog -Message ("ERROR: Invalid ODBC discovery record: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(248, 113, 113))
                }
                continue
            }

            if ($message -match '(?i)^ERROR:|finished with exit code [1-9]') {
                $script:CurrentJob.HadErrors = $true
            }
            Add-ActivityLog -Message $message -Color (Get-LogColor -Message $message)
            continue
        }

        if ($script:CurrentJob.Kind -eq 'LocalAdminDiscovery') {
            $message = [string]$item
            $localAdminMarker = '@@LOCAL_ADMIN_TARGET@@'
            if ($message.StartsWith($localAdminMarker, [System.StringComparison]::Ordinal)) {
                try {
                    $localAdminRecord = $message.Substring($localAdminMarker.Length) |
                        ConvertFrom-Json -ErrorAction Stop
                    [void]$script:CurrentJob.Results.Add([PSCustomObject]@{
                        Server = [string]$localAdminRecord.Server
                        Group  = [string]$localAdminRecord.Group
                        Name   = [string]$localAdminRecord.Name
                        Class  = [string]$localAdminRecord.Class
                        Path   = [string]$localAdminRecord.Path
                    })
                } catch {
                    $script:CurrentJob.HadErrors = $true
                    Add-ActivityLog -Message ("ERROR: Invalid local-admin discovery record: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(248, 113, 113))
                }
                continue
            }

            if ($message -match '(?i)^ERROR:|finished with exit code [1-9]') {
                $script:CurrentJob.HadErrors = $true
            }
            Add-ActivityLog -Message $message -Color (Get-LogColor -Message $message)
            continue
        }

        if ($script:CurrentJob.Kind -eq 'ServiceDiscovery') {
            $message = [string]$item
            $serviceMarker = '@@SERVICE_TARGET@@'
            if ($message.StartsWith($serviceMarker, [System.StringComparison]::Ordinal)) {
                try {
                    $serviceRecord = $message.Substring($serviceMarker.Length) |
                        ConvertFrom-Json -ErrorAction Stop
                    [void]$script:CurrentJob.Results.Add([PSCustomObject]@{
                        TargetType  = [string]$serviceRecord.TargetType
                        Server      = [string]$serviceRecord.Server
                        Name        = [string]$serviceRecord.Name
                        DisplayName = [string]$serviceRecord.DisplayName
                        State       = [string]$serviceRecord.State
                        StartName   = [string]$serviceRecord.StartName
                        FileName    = [string]$serviceRecord.FileName
                        Path        = [string]$serviceRecord.Path
                        Entry       = [string]$serviceRecord.Entry
                        MatchSource = [string]$serviceRecord.MatchSource
                        IsOptInCandidate = [bool]$serviceRecord.IsOptInCandidate
                        MatchReason = [string]$serviceRecord.MatchReason
                    })
                } catch {
                    $script:CurrentJob.HadErrors = $true
                    Add-ActivityLog -Message ("ERROR: Invalid service discovery record: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(248, 113, 113))
                }
                continue
            }

            if ($message -match '(?i)^ERROR:|finished with exit code [1-9]') {
                $script:CurrentJob.HadErrors = $true
            }
            Add-ActivityLog -Message $message -Color (Get-LogColor -Message $message)
            continue
        }

        if ($script:CurrentJob.Kind -eq 'FolderDiscovery') {
            $recordKind = [string]$item.Kind
            if ($recordKind -eq 'Folder') {
                [void]$script:CurrentJob.Results.Add([PSCustomObject]@{
                    Server = [string]$item.Server
                    Name   = [string]$item.Name
                })
                continue
            }

            $message = [string]$item.Message
            if ($message -match '^Listing client folders on (.+)\.\.\.$') {
                $currentServer = $matches[1]
                $serverKey = $currentServer.ToUpperInvariant()
                if (-not $script:CurrentJob.SeenServers.ContainsKey($serverKey)) {
                    $script:CurrentJob.SeenServers[$serverKey] = $true
                }
                $progressServers = @($script:CurrentJob.Context.ProgressServers)
                $serverNumber = @($script:CurrentJob.SeenServers.Keys).Count
                Set-StatusText -Text ('Running: Client folder discovery | Server {0} of {1}: {2}' -f $serverNumber, $progressServers.Count, $currentServer)
                $progressBar.Style = 'Continuous'
                $progressBar.Maximum = [math]::Max(1, $progressServers.Count)
                $progressBar.Value = [math]::Min($progressBar.Maximum, $serverNumber)
            }
            if ($recordKind -eq 'Error') {
                $script:CurrentJob.HadErrors = $true
            }
            Add-ActivityLog -Message $message -Color (Get-LogColor -Message $message)
            continue
        }

        $message = [string]$item
        Add-ActivityLog -Message $message -Color (Get-LogColor -Message $message)
    }

    $errors = $script:CurrentJob.PowerShell.Streams.Error
    while ($script:CurrentJob.ErrorIndex -lt $errors.Count) {
        $errorRecord = $errors[$script:CurrentJob.ErrorIndex]
        $script:CurrentJob.ErrorIndex++
        Add-ActivityLog -Message ("ERROR: {0}" -f $errorRecord.ToString()) -Color ([System.Drawing.Color]::FromArgb(248, 113, 113))
    }
}

$jobTimer = New-Object System.Windows.Forms.Timer
$jobTimer.Interval = 200
$jobTimer.Add_Tick({
    if ($null -eq $script:CurrentJob) {
        $jobTimer.Stop()
        return
    }

    Drain-CurrentJobOutput
    if ($script:CurrentJob.AsyncResult.IsCompleted) {
        $jobTimer.Stop()
        Drain-CurrentJobOutput

        try {
            $script:CurrentJob.PowerShell.EndInvoke($script:CurrentJob.AsyncResult)
        } catch {
            Add-ActivityLog -Message ("ERROR: Background action ended unexpectedly: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(248, 113, 113))
        }
        Drain-CurrentJobOutput

        $completedName = $script:CurrentJob.Name
        $completedKind = $script:CurrentJob.Kind
        $wasCancelled = $script:CurrentJob.Cancelled
        $completedResults = @($script:CurrentJob.Results)
        $hadDiscoveryErrors = [bool]$script:CurrentJob.HadErrors
        $completedContext = $script:CurrentJob.Context
        $completedSteps = @($script:CurrentJob.Steps)
        $completedStepResults = @($script:CurrentJob.StepResults)
        $completedStartedAt = $script:CurrentJob.StartedAt
        $completedSessionId = [string]$script:CurrentJob.SessionId
        $script:CurrentJob.PowerShell.Dispose()
        $script:CurrentJob = $null

        Set-ActionState -Busy $false
        Set-StatusText -Text $(if ($wasCancelled) { 'Cancelled' } else { 'Ready' })
        Add-ActivityLog -Message (
            if ($wasCancelled) {
                "Background action cancelled: $completedName"
            } else {
                "Background action completed: $completedName"
            }
        ) -Color (
            if ($wasCancelled) {
                [System.Drawing.Color]::FromArgb(251, 191, 36)
            } else {
                [System.Drawing.Color]::FromArgb(103, 232, 249)
            }
        )

        if ($completedKind -eq 'WorkerSteps') {
            try {
                $reportPath = Write-OperationReport `
                    -Steps $completedSteps `
                    -StepResults $completedStepResults `
                    -SessionId $completedSessionId `
                    -ActionName $completedName `
                    -StartedAt $completedStartedAt `
                    -EndedAt (Get-Date) `
                    -Cancelled $wasCancelled
                if (-not [string]::IsNullOrWhiteSpace($reportPath)) {
                    Add-ActivityLog -Message ("Operation report: {0}" -f $reportPath) -Color ([System.Drawing.Color]::FromArgb(74, 222, 128))
                }
            } catch {
                Add-ActivityLog -Message ("ERROR: Unable to write operation report: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(248, 113, 113))
            }
        }

        if ($completedKind -eq 'ScanDiscovery' -and -not $wasCancelled) {
            $nameOnlyCandidateCount = @($completedResults | Where-Object { [bool]$_.IsOptInCandidate }).Count
            if ($nameOnlyCandidateCount -gt 0) {
                [void](Show-NameOnlyCandidateSelectionDialog `
                    -ScanResults $completedResults `
                    -Instance ([string]$completedContext.Instance))
                $selectedCandidateCount = @(
                    $completedResults |
                        Where-Object { [bool]$_.IsOptInCandidate -and [bool]$_.OptInSelected }
                ).Count
                Add-ActivityLog -Message ("Name-only candidate review completed: {0} of {1} candidate(s) added to Scan Results." -f $selectedCandidateCount, $nameOnlyCandidateCount) -Color ([System.Drawing.Color]::FromArgb(103, 232, 249))
            }
            $script:LatestScanContext = $completedContext
            Update-ScanResultsPanel `
                -ScanResults $completedResults `
                -Instance ([string]$completedContext.Instance) `
                -HadDiscoveryErrors $hadDiscoveryErrors
            Add-ActivityLog -Message ("Scan displayed {0} result(s). No configuration was changed." -f $completedResults.Count) -Color ([System.Drawing.Color]::FromArgb(103, 232, 249))
        } elseif ($completedKind -eq 'ServiceDiscovery' -and -not $wasCancelled) {
            $candidateCount = @($completedResults | Where-Object { [bool]$_.IsOptInCandidate }).Count
            $completedResults = @(
                $completedResults |
                    Where-Object {
                        -not [bool]$_.IsOptInCandidate -or
                        (Test-OptInCandidateSelected -Category 'Services' -Server ([string]$_.Server) -Name ([string]$_.Name))
                    }
            )
            $selectedCandidateCount = @($completedResults | Where-Object { [bool]$_.IsOptInCandidate }).Count
            if ($candidateCount -gt 0) {
                Add-ActivityLog -Message ("Service discovery found {0} name-only candidate(s); {1} checked candidate(s) passed live revalidation and were included." -f $candidateCount, $selectedCandidateCount) -Color ([System.Drawing.Color]::FromArgb(103, 232, 249))
            }
            if ($hadDiscoveryErrors) {
                [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    'Service or GCS list discovery failed on one or more servers. No deletion was started. Review the background activity log.',
                    'Service discovery incomplete',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
            } elseif ($completedResults.Count -eq 0) {
                $noTargetMessage = if ($candidateCount -gt 0) {
                    "No automatic service targets matched, and none of the $candidateCount name-only candidate(s) were checked in the latest scan."
                } else {
                    "No services or GCS service-list entries matched instance token '$($completedContext.Token)'."
                }
                [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    $noTargetMessage,
                    'No matching service targets',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
            } elseif (Show-ServiceDeletionConfirmation `
                -ServiceTargets $completedResults `
                -AccountToken ([string]$completedContext.Token)) {
                $manifest = ConvertTo-ServiceManifestArgument -ServiceTargets $completedResults
                $steps = New-Object System.Collections.ArrayList
                [void]$steps.Add((New-WorkerStep -Title 'Delete confirmed services and clean GCS lists' -ScriptPath $serviceScript -Arguments @(
                    '-ComputerName', [string]$completedContext.Servers,
                    '-LogOnAsToken', [string]$completedContext.Token,
                    '-TargetManifestBase64', $manifest,
                    '-Force'
                ) -Category 'Services' -Server ([string]$completedContext.Servers) `
                    -ReportTargets @(ConvertTo-ActionReportTargets -Records $completedResults -DefaultCategory 'Services')))
                Start-BackgroundSteps -Steps $steps -ActionName 'Delete services and clean GCS lists' `
                    -Context ([PSCustomObject]@{ ProgressServers = @(ConvertTo-ServerArray -Text @([string]$completedContext.Servers)) })
                if ($null -ne $script:CurrentJob) {
                    $jobTimer.Start()
                }
            } else {
                Add-ActivityLog -Message 'Service deletion cancelled at the confirmation window.' -Color ([System.Drawing.Color]::FromArgb(251, 191, 36))
            }
        } elseif ($completedKind -eq 'IISDiscovery' -and -not $wasCancelled) {
            $candidateCount = @($completedResults | Where-Object { [bool]$_.IsOptInCandidate }).Count
            $completedResults = @(
                $completedResults |
                    Where-Object {
                        -not [bool]$_.IsOptInCandidate -or
                        (Test-OptInCandidateSelected -Category 'ApplicationPools' -Server ([string]$_.Server) -Name ([string]$_.Name))
                    }
            )
            $selectedCandidateCount = @($completedResults | Where-Object { [bool]$_.IsOptInCandidate }).Count
            if ($candidateCount -gt 0) {
                Add-ActivityLog -Message ("IIS discovery found {0} name-only pool candidate(s); {1} checked candidate(s) passed live revalidation and were included." -f $candidateCount, $selectedCandidateCount) -Color ([System.Drawing.Color]::FromArgb(103, 232, 249))
            }
            if ($hadDiscoveryErrors) {
                [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    'IIS discovery failed on one or more servers. No IIS deletion was started. Review the background activity log.',
                    'IIS discovery incomplete',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
            } elseif ($completedResults.Count -eq 0) {
                $noTargetMessage = if ($candidateCount -gt 0) {
                    "No automatic IIS targets matched, and none of the $candidateCount name-only application-pool candidate(s) were checked in the latest scan."
                } else {
                    "No IIS sites, child applications, or application pools matched '$($completedContext.Instance)'."
                }
                [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    $noTargetMessage,
                    'No matching IIS targets',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
            } elseif (Show-IISDeletionConfirmation `
                -IISTargets $completedResults `
                -Instance ([string]$completedContext.Instance) `
                -MatchMode ([string]$completedContext.MatchMode)) {
                $plannedTargets = @($completedResults | Where-Object { [bool]$_.Planned })
                $manifest = ConvertTo-IISTargetManifestArgument -IISTargets $plannedTargets
                $steps = New-Object System.Collections.ArrayList
                [void]$steps.Add((New-WorkerStep -Title 'Delete confirmed IIS targets' -ScriptPath $iisScript -Arguments @(
                    '-ComputerName', [string]$completedContext.Servers,
                    '-Instance', [string]$completedContext.Instance,
                    '-MatchMode', [string]$completedContext.MatchMode,
                    '-TargetManifestBase64', $manifest,
                    '-Force'
                ) -Category 'IIS' -Server ([string]$completedContext.Servers) `
                    -ReportTargets @(ConvertTo-ActionReportTargets -Records $plannedTargets -DefaultCategory 'IIS')))
                Start-BackgroundSteps -Steps $steps -ActionName 'IIS Pool and Site Decom' `
                    -Context ([PSCustomObject]@{ ProgressServers = @(ConvertTo-ServerArray -Text @([string]$completedContext.Servers)) })
                if ($null -ne $script:CurrentJob) {
                    $jobTimer.Start()
                }
            } else {
                Add-ActivityLog -Message 'IIS deletion cancelled at the confirmation window.' -Color ([System.Drawing.Color]::FromArgb(251, 191, 36))
            }
        } elseif ($completedKind -eq 'ODBCDiscovery' -and -not $wasCancelled) {
            if ($hadDiscoveryErrors) {
                [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    'ODBC System DSN discovery failed on one or more servers. No DSN was deleted. Review the background activity log.',
                    'ODBC discovery incomplete',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
            } elseif ($completedResults.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    ("No 32-bit or 64-bit System DSNs matched '{0}'." -f $completedContext.Instance),
                    'No matching System DSNs',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
            } elseif (Show-ODBCCleanupConfirmation `
                -ODBCTargets $completedResults `
                -Instance ([string]$completedContext.Instance)) {
                $manifest = ConvertTo-ODBCTargetManifestArgument -ODBCTargets $completedResults
                $steps = New-Object System.Collections.ArrayList
                [void]$steps.Add((New-WorkerStep -Title 'Delete confirmed ODBC System DSNs' -ScriptPath $odbcScript -Arguments @(
                    '-ComputerName', [string]$completedContext.Servers,
                    '-Instance', [string]$completedContext.Instance,
                    '-TargetManifestBase64', $manifest,
                    '-Force'
                ) -Category 'ODBCDataSources' -Server ([string]$completedContext.Servers) `
                    -ReportTargets @(ConvertTo-ActionReportTargets -Records $completedResults -DefaultCategory 'ODBCDataSources')))
                Start-BackgroundSteps -Steps $steps -ActionName 'ODBC System DSN Cleanup' `
                    -Context ([PSCustomObject]@{ ProgressServers = @(ConvertTo-ServerArray -Text @([string]$completedContext.Servers)) })
                if ($null -ne $script:CurrentJob) {
                    $jobTimer.Start()
                }
            } else {
                Add-ActivityLog -Message 'ODBC System DSN cleanup cancelled at the confirmation window.' -Color ([System.Drawing.Color]::FromArgb(251, 191, 36))
            }
        } elseif ($completedKind -eq 'LocalAdminDiscovery' -and -not $wasCancelled) {
            if ($hadDiscoveryErrors) {
                [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    'Local-admin discovery failed on one or more servers. No group membership was changed. Review the background activity log.',
                    'Local-admin discovery incomplete',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
            } elseif ($completedResults.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    $form,
                    ("No local Administrators group principals matched '{0}'." -f $completedContext.Instance),
                    'No matching local-admin targets',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
            } elseif (Show-LocalAdminCleanupConfirmation `
                -LocalAdminTargets $completedResults `
                -Instance ([string]$completedContext.Instance)) {
                $manifest = ConvertTo-LocalAdminTargetManifestArgument -LocalAdminTargets $completedResults
                $steps = New-Object System.Collections.ArrayList
                [void]$steps.Add((New-WorkerStep -Title 'Remove confirmed local Administrators groups' -ScriptPath $localAdminScript -Arguments @(
                    '-ComputerName', [string]$completedContext.Servers,
                    '-Instance', [string]$completedContext.Instance,
                    '-TargetManifestBase64', $manifest,
                    '-Force'
                ) -Category 'LocalAdmins' -Server ([string]$completedContext.Servers) `
                    -ReportTargets @(ConvertTo-ActionReportTargets -Records $completedResults -DefaultCategory 'LocalAdmins')))
                Start-BackgroundSteps -Steps $steps -ActionName 'Local Admin Cleanup' `
                    -Context ([PSCustomObject]@{ ProgressServers = @(ConvertTo-ServerArray -Text @([string]$completedContext.Servers)) })
                if ($null -ne $script:CurrentJob) {
                    $jobTimer.Start()
                }
            } else {
                Add-ActivityLog -Message 'Local-admin cleanup cancelled at the confirmation window.' -Color ([System.Drawing.Color]::FromArgb(251, 191, 36))
            }
        } elseif ($completedKind -eq 'FolderDiscovery' -and -not $wasCancelled) {
            if ($completedResults.Count -eq 0) {
                $message = if ($hadDiscoveryErrors) {
                    'No folders could be listed. Review the background activity log for server errors.'
                } else {
                    'No directories were found under M:\OBOL\Clients on the selected servers.'
                }
                Show-InputError $message
            } else {
                $selectedTargets = @(
                    Show-FolderSelectionDialog `
                        -FolderTargets $completedResults `
                        -HadDiscoveryErrors $hadDiscoveryErrors `
                        -InstanceToken ($instanceTextBox.Text.Trim())
                )
                if ($selectedTargets.Count -gt 0) {
                    $script:SelectedFolderTargets = @($selectedTargets)
                    Update-FolderSelectionSummary
                    Add-ActivityLog -Message ("Selected {0} server/folder target(s) for folder deletion." -f $selectedTargets.Count) -Color ([System.Drawing.Color]::FromArgb(103, 232, 249))
                }
            }
        }
    }
})

$serviceServersTextBox.Add_TextChanged({
    Clear-FolderSelection -WriteLog
    Clear-OptInCandidateSelections -WriteLog
})

$applicationServersTextBox.Add_TextChanged({
    Clear-FolderSelection -WriteLog
    Clear-OptInCandidateSelections -WriteLog
})

$instanceTextBox.Add_TextChanged({
    Clear-OptInCandidateSelections -WriteLog
    $matchDescription = Get-ServiceAccountMatchDescription
    $serviceAccountTextBox.Text = if ([string]::IsNullOrWhiteSpace($matchDescription)) {
        'Account name contains instance token'
    } else {
        $matchDescription
    }
})

$matchModeComboBox.Add_SelectedIndexChanged({
    Clear-OptInCandidateSelections -WriteLog
})

$selectFoldersButton.Add_Click({
    $servers = @(Get-CombinedServerArray)
    if ($servers.Count -eq 0) {
        Show-InputError 'Enter at least one service or application server before choosing folders.'
        return
    }

    foreach ($server in $servers) {
        if ($server -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
            Show-InputError ("Invalid server name '{0}'." -f $server)
            return
        }
    }

    Start-FolderDiscovery -Servers $servers
    if ($null -ne $script:CurrentJob) {
        $jobTimer.Start()
    }
})

$deleteServicesButton.Add_Click({
    $servers = ConvertTo-ArgumentText -Text $serviceServersTextBox.Text
    $instance = $instanceTextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($servers)) {
        Show-InputError 'Enter at least one service server.'
        return
    }
    if ([string]::IsNullOrWhiteSpace($instance)) {
        Show-InputError "Enter the instance token to find in each service's Log On As account name."
        return
    }
    if ($instance.Length -lt 3 -or $instance -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        Show-InputError 'The service instance token must be at least 3 characters and contain only letters, numbers, period, underscore, or hyphen.'
        return
    }

    $steps = New-Object System.Collections.ArrayList
    [void]$steps.Add((New-WorkerStep -Title 'Discover matching services' -ScriptPath $serviceScript -Arguments @(
        '-ComputerName', $servers,
        '-LogOnAsToken', $instance,
        '-DiscoveryOnly',
        '-EmitDiscoveryJson',
        '-IncludeNameOnlyCandidates'
    )))
    $context = [PSCustomObject]@{
        Servers = $servers
        Token   = $instance
        ProgressServers = @(ConvertTo-ServerArray -Text @($serviceServersTextBox.Text))
    }
    Start-BackgroundSteps `
        -Steps $steps `
        -ActionName 'Service discovery' `
        -JobKind 'ServiceDiscovery' `
        -Context $context
    if ($null -ne $script:CurrentJob) {
        $jobTimer.Start()
    }
})

$iisButton.Add_Click({
    $servers = ConvertTo-ArgumentText -Text $applicationServersTextBox.Text
    $instance = $instanceTextBox.Text.Trim()
    $matchMode = [string]$matchModeComboBox.SelectedItem
    if ([string]::IsNullOrWhiteSpace($servers)) {
        Show-InputError 'Enter at least one application / IIS server.'
        return
    }
    if ([string]::IsNullOrWhiteSpace($instance)) {
        Show-InputError 'Enter the IIS instance token.'
        return
    }

    $steps = New-Object System.Collections.ArrayList
    [void]$steps.Add((New-WorkerStep -Title 'Discover matching IIS targets' -ScriptPath $iisScript -Arguments @(
        '-ComputerName', $servers,
        '-Instance', $instance,
        '-MatchMode', $matchMode,
        '-DiscoveryOnly',
        '-EmitDiscoveryJson'
    )))
    $context = [PSCustomObject]@{
        Servers   = $servers
        Instance  = $instance
        MatchMode = $matchMode
        ProgressServers = @(ConvertTo-ServerArray -Text @($applicationServersTextBox.Text))
    }
    Start-BackgroundSteps `
        -Steps $steps `
        -ActionName 'IIS discovery' `
        -JobKind 'IISDiscovery' `
        -Context $context
    if ($null -ne $script:CurrentJob) {
        $jobTimer.Start()
    }
})

$localAdminButton.Add_Click({
    $servers = @(Get-CombinedServerArray)
    $instance = $instanceTextBox.Text.Trim()

    if ($servers.Count -eq 0) {
        Show-InputError 'Enter at least one service or application server.'
        return
    }
    if ([string]::IsNullOrWhiteSpace($instance)) {
        Show-InputError 'Enter the instance token.'
        return
    }
    foreach ($server in $servers) {
        if ($server -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
            Show-InputError ("Invalid server name '{0}'." -f $server)
            return
        }
    }

    $serverText = $servers -join ','
    $steps = New-Object System.Collections.ArrayList
    [void]$steps.Add((New-WorkerStep -Title 'Discover matching local Administrators groups' -ScriptPath $localAdminScript -Arguments @(
        '-ComputerName', $serverText,
        '-Instance', $instance,
        '-DiscoveryOnly',
        '-EmitDiscoveryJson'
    )))
    $context = [PSCustomObject]@{
        Servers  = $serverText
        Instance = $instance
        ProgressServers = @($servers)
    }
    Start-BackgroundSteps `
        -Steps $steps `
        -ActionName 'Local-admin discovery' `
        -JobKind 'LocalAdminDiscovery' `
        -Context $context
    if ($null -ne $script:CurrentJob) {
        $jobTimer.Start()
    }
})

$odbcButton.Add_Click({
    $servers = @(Get-CombinedServerArray)
    $instance = $instanceTextBox.Text.Trim()

    if ($servers.Count -eq 0) {
        Show-InputError 'Enter at least one service or application server.'
        return
    }
    if ([string]::IsNullOrWhiteSpace($instance)) {
        Show-InputError 'Enter the instance token used to match System DSN names.'
        return
    }
    if ($instance.Length -lt 3 -or $instance -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        Show-InputError 'The ODBC instance token must be at least 3 characters and contain only letters, numbers, period, underscore, or hyphen.'
        return
    }
    foreach ($server in $servers) {
        if ($server -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
            Show-InputError ("Invalid server name '{0}'." -f $server)
            return
        }
    }

    $serverText = $servers -join ','
    $steps = New-Object System.Collections.ArrayList
    [void]$steps.Add((New-WorkerStep -Title 'Discover matching ODBC System DSNs' -ScriptPath $odbcScript -Arguments @(
        '-ComputerName', $serverText,
        '-Instance', $instance,
        '-DiscoveryOnly',
        '-EmitDiscoveryJson'
    )))
    $context = [PSCustomObject]@{
        Servers  = $serverText
        Instance = $instance
        ProgressServers = @($servers)
    }
    Start-BackgroundSteps `
        -Steps $steps `
        -ActionName 'ODBC System DSN discovery' `
        -JobKind 'ODBCDiscovery' `
        -Context $context
    if ($null -ne $script:CurrentJob) {
        $jobTimer.Start()
    }
})

$folderButton.Add_Click({
    $combinedServers = @(Get-CombinedServerArray)
    $folderTargets = @($script:SelectedFolderTargets)

    if ($combinedServers.Count -eq 0) {
        Show-InputError 'Enter at least one service or application server.'
        return
    }
    if ($folderTargets.Count -eq 0) {
        Show-InputError 'Choose client folders from the directory list first.'
        return
    }

    if (-not (Confirm-FolderDeletion -FolderTargets $folderTargets)) {
        return
    }

    $steps = New-Object System.Collections.ArrayList
    foreach ($serverGroup in @($folderTargets | Group-Object Server)) {
        $folderNames = @($serverGroup.Group | ForEach-Object { $_.Name } | Select-Object -Unique)
        $encodedFolders = ConvertTo-EncodedFolderArgument -FolderNames $folderNames
        [void]$steps.Add((New-WorkerStep -Title ("Folder deletion - {0}" -f $serverGroup.Name) -ScriptPath $folderScript -Arguments @(
            '-ComputerName', [string]$serverGroup.Name,
            '-EncodedFolderName', $encodedFolders,
            '-Force'
        ) -Category 'Folders' -Server ([string]$serverGroup.Name) `
            -ReportTargets @(ConvertTo-ActionReportTargets -Records @($serverGroup.Group) -DefaultCategory 'Folders')))
    }

    Start-BackgroundSteps -Steps $steps -ActionName 'Folder deletion' `
        -Context ([PSCustomObject]@{ ProgressServers = @($folderTargets | ForEach-Object { [string]$_.Server } | Select-Object -Unique) })
    $jobTimer.Start()
})

$scanButton.Add_Click({
    $serviceServers = ConvertTo-ArgumentText -Text $serviceServersTextBox.Text
    $applicationServers = ConvertTo-ArgumentText -Text $applicationServersTextBox.Text
    $combinedServers = @(Get-CombinedServerArray)
    $instance = $instanceTextBox.Text.Trim()
    $matchMode = [string]$matchModeComboBox.SelectedItem

    if ($combinedServers.Count -eq 0) {
        Show-InputError 'Enter at least one service or application server before scanning.'
        return
    }
    if ([string]::IsNullOrWhiteSpace($instance)) {
        Show-InputError 'Enter the instance token before scanning.'
        return
    }
    if ($instance.Length -lt 3 -or $instance -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        Show-InputError 'The scan instance token must be at least 3 characters and contain only letters, numbers, period, underscore, or hyphen.'
        return
    }
    foreach ($server in $combinedServers) {
        if ($server -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
            Show-InputError ("Invalid server name '{0}'." -f $server)
            return
        }
    }

    $steps = New-Object System.Collections.ArrayList
    [void]$steps.Add((New-WorkerStep -Title 'Read-only inventory scan' -ScriptPath $scanScript -Arguments @(
        '-ServiceComputerName', $serviceServers,
        '-ApplicationComputerName', $applicationServers,
        '-Instance', $instance,
        '-IISMatchMode', $matchMode
    )))
    $context = [PSCustomObject]@{
        Instance           = $instance
        ServiceServers     = @(ConvertTo-ServerArray -Text @($serviceServersTextBox.Text))
        ApplicationServers = @(ConvertTo-ServerArray -Text @($applicationServersTextBox.Text))
        ProgressServers    = @($combinedServers)
        MatchMode          = $matchMode
    }
    Start-BackgroundSteps `
        -Steps $steps `
        -ActionName 'Read-only inventory scan' `
        -JobKind 'ScanDiscovery' `
        -Context $context
    if ($null -ne $script:CurrentJob) {
        $jobTimer.Start()
    }
})

$whatIfButton.Add_Click({
    $steps = New-Object System.Collections.ArrayList
    $instance = $instanceTextBox.Text.Trim()

    $serviceServers = ConvertTo-ArgumentText -Text $serviceServersTextBox.Text
    if (-not [string]::IsNullOrWhiteSpace($serviceServers)) {
        if ([string]::IsNullOrWhiteSpace($instance)) {
            Show-InputError "Service preview requires an instance token to find in each Log On As account name."
            return
        }
        if ($instance.Length -lt 3 -or $instance -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
            Show-InputError 'The service instance token must be at least 3 characters and contain only letters, numbers, period, underscore, or hyphen.'
            return
        }
        [void]$steps.Add((New-WorkerStep -Title 'WhatIf - Services' -ScriptPath $serviceScript -Arguments @(
            '-ComputerName', $serviceServers,
            '-LogOnAsToken', $instance,
            '-WhatIf',
            '-Force'
        )))
    }

    $applicationServers = ConvertTo-ArgumentText -Text $applicationServersTextBox.Text
    $matchMode = [string]$matchModeComboBox.SelectedItem
    if (-not [string]::IsNullOrWhiteSpace($applicationServers)) {
        if ([string]::IsNullOrWhiteSpace($instance)) {
            Show-InputError 'IIS preview requires an instance token.'
            return
        }
        [void]$steps.Add((New-WorkerStep -Title 'WhatIf - IIS' -ScriptPath $iisScript -Arguments @(
            '-ComputerName', $applicationServers,
            '-Instance', $instance,
            '-MatchMode', $matchMode,
            '-WhatIf',
            '-Force'
        )))
    }

    $folderTargets = @($script:SelectedFolderTargets)
    $combinedServers = @(Get-CombinedServerArray)
    if ($combinedServers.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($instance)) {
        [void]$steps.Add((New-WorkerStep -Title 'WhatIf - Local Admin Cleanup' -ScriptPath $localAdminScript -Arguments @(
            '-ComputerName', ($combinedServers -join ','),
            '-Instance', $instance,
            '-WhatIf',
            '-Force'
        )))
        [void]$steps.Add((New-WorkerStep -Title 'WhatIf - ODBC System DSNs' -ScriptPath $odbcScript -Arguments @(
            '-ComputerName', ($combinedServers -join ','),
            '-Instance', $instance,
            '-WhatIf',
            '-Force'
        )))
    }

    if ($folderTargets.Count -gt 0) {
        if ($combinedServers.Count -eq 0) {
            Show-InputError 'Folder preview requires at least one service or application server.'
            return
        }

        foreach ($serverGroup in @($folderTargets | Group-Object Server)) {
            $folderNames = @($serverGroup.Group | ForEach-Object { $_.Name } | Select-Object -Unique)
            $encodedFolders = ConvertTo-EncodedFolderArgument -FolderNames $folderNames
            [void]$steps.Add((New-WorkerStep -Title ("WhatIf - Folders - {0}" -f $serverGroup.Name) -ScriptPath $folderScript -Arguments @(
                '-ComputerName', [string]$serverGroup.Name,
                '-EncodedFolderName', $encodedFolders,
                '-WhatIf',
                '-Force'
            )))
        }
    }

    if ($steps.Count -eq 0) {
        Show-InputError 'Enter inputs for at least one action before running What If.'
        return
    }

    Start-BackgroundSteps -Steps $steps -ActionName 'What If' `
        -Context ([PSCustomObject]@{ ProgressServers = @(Get-CombinedServerArray) })
    $jobTimer.Start()
})

$cancelButton.Add_Click({
    if ($null -eq $script:CurrentJob) { return }

    $choice = [System.Windows.Forms.MessageBox]::Show(
        'Stop the current background action?',
        'Cancel action',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    )
    if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
        $script:CurrentJob.Cancelled = $true
        try {
            $script:CurrentJob.PowerShell.Stop()
        } catch {
            Add-ActivityLog -Message ("WARNING: Unable to stop immediately: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(251, 191, 36))
        }
    }
})

$clearLogButton.Add_Click({
    $activityLog.Clear()
    if (-not [string]::IsNullOrWhiteSpace($script:ActivityLogPath)) {
        Add-ActivityLog -Message ("On-screen activity cleared. The session text log was retained: {0}" -f $script:ActivityLogPath)
    } else {
        Add-ActivityLog -Message 'On-screen activity cleared.'
    }
})

$exportScanButton.Add_Click({
    if (@($script:LatestVisibleScanResults).Count -eq 0) {
        Show-InputError 'Run SCAN and wait for results before exporting.'
        return
    }

    $safeInstance = if ([string]::IsNullOrWhiteSpace($script:LatestScanInstance)) {
        'All'
    } else {
        $script:LatestScanInstance -replace '[^A-Za-z0-9._-]', '_'
    }
    $fileTimestamp = if ($script:LatestScanTimestamp -eq [datetime]::MinValue) {
        Get-Date -Format 'yyyyMMdd_HHmmss'
    } else {
        $script:LatestScanTimestamp.ToString('yyyyMMdd_HHmmss')
    }

    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Title = 'Export Scan Results'
    $saveDialog.Filter = 'Excel workbooks (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    $saveDialog.DefaultExt = 'xlsx'
    $saveDialog.AddExtension = $true
    $saveDialog.OverwritePrompt = $true
    $saveDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    $saveDialog.FileName = 'POST_DECOM_Scan_{0}_{1}.xlsx' -f $safeInstance, $fileTimestamp

    try {
        if ($saveDialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }

        $exportResult = Export-ScanResultsWorkbook `
            -Path $saveDialog.FileName `
            -ScanResults $script:LatestVisibleScanResults `
            -Instance $script:LatestScanInstance `
            -ScanTimestamp $script:LatestScanTimestamp

        Add-ActivityLog -Message ("Exported {0} scan result(s) to {1} Excel worksheet(s): {2}" -f $exportResult.RowCount, $exportResult.SheetCount, $exportResult.Path) -Color ([System.Drawing.Color]::FromArgb(74, 222, 128))
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            ("Exported {0} scan result(s) into {1} worksheets:`r`nServices, IIS Sites, Application Pools, ODBC Data Sources, Local Admins, and Client Folders.`r`n`r`n{2}" -f $exportResult.RowCount, $exportResult.SheetCount, $exportResult.Path),
            'Scan export complete',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    } catch {
        Add-ActivityLog -Message ("ERROR: Unable to export scan results: {0}" -f $_.Exception.Message) -Color ([System.Drawing.Color]::FromArgb(248, 113, 113))
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            ("Unable to export scan results.`r`n`r`n{0}" -f $_.Exception.Message),
            'Scan export failed',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $saveDialog.Dispose()
    }
})

$form.Add_FormClosing({
    param($sender, $eventArgs)

    if ($null -eq $script:CurrentJob) { return }

    $choice = [System.Windows.Forms.MessageBox]::Show(
        'A background action is still running. Stop it and close the application?',
        'Close application',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    )
    if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
        $eventArgs.Cancel = $true
        return
    }

    try {
        $script:CurrentJob.PowerShell.Stop()
        $script:CurrentJob.PowerShell.Dispose()
    } catch {}
    $script:CurrentJob = $null
})

$form.Add_FormClosed({
    if (-not [string]::IsNullOrWhiteSpace($script:ActivityLogPath) -and
        -not $script:ActivityLogWriteFailed) {
        try {
            Write-ActivityLogFileLine -Path $script:ActivityLogPath -Line ('[{0:HH:mm:ss}] Application closed.' -f (Get-Date))
        } catch {}
    }
})

$renderPaths = @(
    $RenderPreviewPath
    $RenderFolderPickerPreviewPath
    $RenderServiceConfirmationPreviewPath
    $RenderIISConfirmationPreviewPath
    $RenderLocalAdminConfirmationPreviewPath
    $RenderScanResultsPreviewPath
    $RenderCandidateSelectionPreviewPath
    $RenderODBCConfirmationPreviewPath
)
$hasRenderPath = @($renderPaths | Where-Object {
    -not [string]::IsNullOrWhiteSpace([string]$_)
}).Count -gt 0
$isRenderMode = $hasRenderPath -or [bool]$RenderScanResultsWidePreview

if (-not $isRenderMode) {
    try {
        $script:ActivityLogPath = New-ActivityLogFile `
            -Directory (Join-Path $PSScriptRoot 'Logs') `
            -SessionId $script:ActivityLogSessionId
    } catch {
        $script:ActivityLogPath = ''
        [System.Windows.Forms.MessageBox]::Show(
            "The automatic Background activity text log could not be created.`r`n`r`n$($_.Exception.Message)`r`n`r`nThe tool can continue, but this session will only be shown on screen.",
            'Background activity log unavailable',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
}

Add-ActivityLog -Message 'Application ready. Use What If before running destructive actions.' -Color ([System.Drawing.Color]::FromArgb(103, 232, 249))
if (-not [string]::IsNullOrWhiteSpace($script:ActivityLogPath)) {
    Add-ActivityLog -Message ("Background activity is being saved automatically to: {0}" -f $script:ActivityLogPath) -Color ([System.Drawing.Color]::FromArgb(103, 232, 249))
}

Update-ScanResultsPanel -ScanResults @() -Instance '' -HadDiscoveryErrors $false

if (-not [string]::IsNullOrWhiteSpace($RenderCandidateSelectionPreviewPath)) {
    [void](Show-NameOnlyCandidateSelectionDialog -ScanResults @(
        [PSCustomObject]@{
            Category = 'Services'; Server = 'service-server-02'; TargetType = 'Service'
            Name = 'IM35880Maintenance'; DisplayName = 'IM35880 Maintenance'; State = 'Running'
            StartName = 'LocalSystem'; MatchReason = 'Service name contains the instance token; Log On As does not.'
            IsOptInCandidate = $true; OptInSelected = $false
        }
        [PSCustomObject]@{
            Category = 'Services'; Server = 'service-server-03'; TargetType = 'Service'
            Name = 'IM35880LegacyWorker'; DisplayName = 'Legacy Worker'; State = 'Stopped'
            StartName = 'OLYMPUS\SharedService'; MatchReason = 'Service name contains the instance token; Log On As does not.'
            IsOptInCandidate = $true; OptInSelected = $false
        }
        [PSCustomObject]@{
            Category = 'ApplicationPools'; Server = 'web-server-01'; TargetType = 'AppPool'
            Name = 'IM35880 Legacy'; State = 'Stopped'; PoolIdentity = 'ApplicationPoolIdentity'
            RemovalStatus = 'Eligible'; ScanDetails = 'Application-pool name contains the instance token; pool identity does not.'
            IsOptInCandidate = $true; OptInSelected = $false
        }
    ) -Instance 'IM35880' -RenderPreviewPath $RenderCandidateSelectionPreviewPath)
    $form.Dispose()
    return
}

if (-not [string]::IsNullOrWhiteSpace($RenderLocalAdminConfirmationPreviewPath)) {
    [void](Show-LocalAdminCleanupConfirmation -LocalAdminTargets @(
        [PSCustomObject]@{
            Server = 'service-server-01'; Group = 'Administrators'; Name = 'apIM35880-Admins'
            Class = 'Group'; Path = 'WinNT://OLYMPUS/apIM35880-Admins,group'
        }
        [PSCustomObject]@{
            Server = 'application-server-01'; Group = 'Administrators'; Name = 'IM35880-Support'
            Class = 'Group'; Path = 'WinNT://OLYMPUS/IM35880-Support,group'
        }
    ) -Instance 'IM35880' -RenderPreviewPath $RenderLocalAdminConfirmationPreviewPath)
    $form.Dispose()
    return
}

if (-not [string]::IsNullOrWhiteSpace($RenderIISConfirmationPreviewPath)) {
    [void](Show-IISDeletionConfirmation -IISTargets @(
        [PSCustomObject]@{
            Server = 'application-server-01'; TargetType = 'Site'; Name = 'IM35880 Int Site'
            SiteName = 'IM35880 Int Site'; ApplicationPath = ''; AppPoolName = ''; State = 'Started'
            Planned = $true; Details = 'Path=M:\OBOL\Clients\Example'; BlockedReason = ''
        }
        [PSCustomObject]@{
            Server = 'web-server-01'; TargetType = 'Application'; Name = 'Iron Mountain Site/IM35880/Downloads'
            SiteName = 'Iron Mountain Site'; ApplicationPath = '/IM35880/Downloads'; AppPoolName = 'IM35880 Downloads'; State = ''
            Planned = $true; Details = 'Physical path=M:\OBOL\Clients\Example\Downloads'; BlockedReason = ''
        }
        [PSCustomObject]@{
            Server = 'web-server-01'; TargetType = 'AppPool'; Name = 'IM35880 Downloads'
            SiteName = ''; ApplicationPath = ''; AppPoolName = 'IM35880 Downloads'; State = 'Started'
            Planned = $true; Details = 'Consumers=Iron Mountain Site/IM35880/Downloads'; BlockedReason = ''
        }
        [PSCustomObject]@{
            Server = 'web-server-01'; TargetType = 'AppPool'; Name = 'IM35880 Shared'
            SiteName = ''; ApplicationPath = ''; AppPoolName = 'IM35880 Shared'; State = 'Started'
            Planned = $false; Details = ''; BlockedReason = 'Used outside removal plan: Iron Mountain Site/OtherApplication'
        }
    ) -Instance 'IM35880' -MatchMode 'Token' -RenderPreviewPath $RenderIISConfirmationPreviewPath)
    $form.Dispose()
    return
}

if (-not [string]::IsNullOrWhiteSpace($RenderScanResultsPreviewPath)) {
    Update-ScanResultsPanel -ScanResults @(
        [PSCustomObject]@{
            Category = 'Services'; Server = 'service-server-01'; TargetType = 'Service'
            Name = 'APIM35880Worker'; DisplayName = 'IM35880 Worker Service'; State = 'Running'
            StartName = 'APIM35880@olympus.gaia.kosmos'; Path = ''; Planned = $true; MatchesToken = $true
            MatchSource = 'Account'; MatchReason = 'Log On As contains the instance token.'; IsOptInCandidate = $false
        }
        [PSCustomObject]@{
            Category = 'Services'; Server = 'service-server-01'; TargetType = 'GCS list entry'
            Name = 'APIM35880Worker'; DisplayName = 'MonitoredServices.txt'; State = 'Remove line'
            StartName = ''; Path = '\\service-server-01\C$\OBOL\Utilities\GCSserviceStartRestarter\MonitoredServices.txt'; Planned = $true; MatchesToken = $true
            MatchSource = 'GCSList'; IsOptInCandidate = $false
        }
        [PSCustomObject]@{
            Category = 'Services'; Server = 'service-server-02'; TargetType = 'Service'
            Name = 'IM35880Maintenance'; DisplayName = 'IM35880 Maintenance'; State = 'Running'
            StartName = 'LocalSystem'; Path = ''; Planned = $true; MatchesToken = $true
            MatchSource = 'NameOnly'; MatchReason = 'Service name contains the instance token; Log On As does not.'
            IsOptInCandidate = $true; OptInSelected = $true
        }
        [PSCustomObject]@{
            Category = 'IISSites'; Server = 'application-server-01'; TargetType = 'Site'
            Name = 'IM35880 Internal Site'; State = 'Started'; Details = 'Bindings=http/*:80'; Planned = $true; MatchesToken = $true
        }
        [PSCustomObject]@{
            Category = 'ApplicationPools'; Server = 'application-server-01'; TargetType = 'AppPool'
            Name = 'IM35880 Downloads'; State = 'Started'; RemovalStatus = 'Eligible'; ScanDetails = 'Consumers=Main Site/IM35880/Downloads'; Planned = $true; MatchesToken = $true
            PoolIdentity = 'OLYMPUS\APIM35880'; MatchSource = 'Identity'; IsOptInCandidate = $false
        }
        [PSCustomObject]@{
            Category = 'ApplicationPools'; Server = 'application-server-01'; TargetType = 'AppPool'
            Name = 'IM35880 Shared'; State = 'Started'; RemovalStatus = 'Blocked'; ScanDetails = 'Used outside removal plan'; Planned = $false; MatchesToken = $true
            PoolIdentity = 'OLYMPUS\APIM35880'; MatchSource = 'Identity'; IsOptInCandidate = $false
        }
        [PSCustomObject]@{
            Category = 'ApplicationPools'; Server = 'web-server-01'; TargetType = 'AppPool'
            Name = 'IM35880 Legacy'; State = 'Stopped'; RemovalStatus = 'Eligible'
            ScanDetails = 'Application-pool name contains the instance token; pool identity does not.'
            Planned = $true; MatchesToken = $true; PoolIdentity = 'ApplicationPoolIdentity'
            MatchSource = 'NameOnly'; MatchReason = 'Application-pool name contains the instance token; pool identity does not.'
            IsOptInCandidate = $true; OptInSelected = $true
        }
        [PSCustomObject]@{
            Category = 'ODBCDataSources'; Server = 'application-server-01'; TargetType = 'System DSN'
            Name = 'IM35880 Reporting'; Architecture = '64-bit'; Driver = 'ODBC Driver 18 for SQL Server'
            DataServer = 'sql-server-01'; Database = 'IM35880'; RegistryPath = 'HKLM\SOFTWARE\ODBC\ODBC.INI\IM35880 Reporting'
            Planned = $true; MatchesToken = $true
        }
        [PSCustomObject]@{
            Category = 'ODBCDataSources'; Server = 'service-server-01'; TargetType = 'System DSN'
            Name = 'IM35880 Legacy'; Architecture = '32-bit'; Driver = 'SQL Server'
            DataServer = 'sql-server-01'; Database = 'IM35880Legacy'; RegistryPath = 'HKLM\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\IM35880 Legacy'
            Planned = $true; MatchesToken = $true
        }
        [PSCustomObject]@{
            Category = 'LocalAdmins'; Server = 'service-server-01'; TargetType = 'Group'
            Name = 'OLYMPUS\IM35880-Admins'; DisplayName = 'Administrators'; Path = 'WinNT://OLYMPUS/IM35880-Admins,group'; Planned = $true; MatchesToken = $true
        }
        [PSCustomObject]@{
            Category = 'Folders'; Server = 'application-server-01'; TargetType = 'Folder'
            Name = 'IM35880 Client'; MatchText = 'Yes'; Path = '\\application-server-01\m$\OBOL\Clients\IM35880 Client'; Planned = $false; MatchesToken = $true
        }
        [PSCustomObject]@{
            Category = 'Folders'; Server = 'application-server-01'; TargetType = 'Folder'
            Name = 'Other Client'; MatchText = 'No'; Path = '\\application-server-01\m$\OBOL\Clients\Other Client'; Planned = $false; MatchesToken = $false
        }
    ) -Instance 'IM35880' -HadDiscoveryErrors $false -RenderPreviewPath $RenderScanResultsPreviewPath
    $form.Dispose()
    return
}

if (-not [string]::IsNullOrWhiteSpace($RenderODBCConfirmationPreviewPath)) {
    [void](Show-ODBCCleanupConfirmation -ODBCTargets @(
        [PSCustomObject]@{
            Server = 'application-server-01'; Architecture = '64-bit'; Name = 'IM35880 Reporting'
            Driver = 'ODBC Driver 18 for SQL Server'; DataServer = 'sql-server-01'; Database = 'IM35880'
            RegistryPath = 'HKLM\SOFTWARE\ODBC\ODBC.INI\IM35880 Reporting'
        }
        [PSCustomObject]@{
            Server = 'service-server-01'; Architecture = '32-bit'; Name = 'IM35880 Legacy'
            Driver = 'SQL Server'; DataServer = 'sql-server-01'; Database = 'IM35880Legacy'
            RegistryPath = 'HKLM\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\IM35880 Legacy'
        }
    ) -Instance 'IM35880' -RenderPreviewPath $RenderODBCConfirmationPreviewPath)
    $form.Dispose()
    return
}

if (-not [string]::IsNullOrWhiteSpace($RenderServiceConfirmationPreviewPath)) {
    [void](Show-ServiceDeletionConfirmation -ServiceTargets @(
        [PSCustomObject]@{
            TargetType = 'Service'
            Server = 'service-server-01'
            Name = 'ExampleService'
            DisplayName = 'Example Application Service'
            State = 'Running'
            StartName = 'Olympus\APIM35880'
        }
        [PSCustomObject]@{
            TargetType = 'Service'
            Server = 'service-server-02'
            Name = 'ExampleWorker'
            DisplayName = 'Example Background Worker'
            State = 'Stopped'
            StartName = 'APIM35880@olympus.gaia.kosmos'
        }
        [PSCustomObject]@{
            TargetType = 'GCSListEntry'
            Server = 'service-server-01'
            Name = 'APIM35880 Worker'
            DisplayName = 'MonitoredServices.txt'
            State = 'Remove line'
            StartName = ''
            FileName = 'MonitoredServices.txt'
            Path = '\\service-server-01\C$\OBOL\Utilities\GCSserviceStartRestarter\MonitoredServices.txt'
            Entry = 'APIM35880 Worker'
        }
        [PSCustomObject]@{
            TargetType = 'GCSListEntry'
            Server = 'service-server-02'
            Name = 'Olympus IM35880 Restarter'
            DisplayName = 'ServicestoRestart.txt'
            State = 'Remove line'
            StartName = ''
            FileName = 'ServicestoRestart.txt'
            Path = '\\service-server-02\C$\OBOL\Utilities\GCSserviceStartRestarter\ServicestoRestart.txt'
            Entry = 'Olympus IM35880 Restarter'
        }
    ) -AccountToken 'IM35880' -RenderPreviewPath $RenderServiceConfirmationPreviewPath)
    $form.Dispose()
    return
}

if (-not [string]::IsNullOrWhiteSpace($RenderFolderPickerPreviewPath)) {
    [void](Show-FolderSelectionDialog -FolderTargets @(
        [PSCustomObject]@{ Server = 'service-server-01'; Name = 'Client-A' }
        [PSCustomObject]@{ Server = 'service-server-01'; Name = 'Client-B' }
        [PSCustomObject]@{ Server = 'application-server-01'; Name = 'Client-A' }
        [PSCustomObject]@{ Server = 'application-server-01'; Name = 'Client-C' }
    ) -InstanceToken 'Client-A' -RenderPreviewPath $RenderFolderPickerPreviewPath)
    $form.Dispose()
    return
}

if (-not [string]::IsNullOrWhiteSpace($RenderPreviewPath)) {
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()

    $bitmap = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
    try {
        $form.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, $form.Width, $form.Height)))
        $bitmap.Save($RenderPreviewPath, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $bitmap.Dispose()
        $form.Close()
        $form.Dispose()
    }
    return
}

[void]$form.ShowDialog()
