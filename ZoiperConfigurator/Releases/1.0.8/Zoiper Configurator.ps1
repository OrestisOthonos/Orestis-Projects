# Zoiper 5 Setup Helper
# This script prompts the user for Zoiper 5 credentials.

$ScriptVersion = '1.0.7'

# --- Self-update configuration ---
# GitHub repo info for update discovery
$GitHubOwner = 'OrestisOthonos'
$GitHubRepo = 'Orestis-Projects'
$GitHubBranch = 'main'
$GitHubReleasesPath = 'ZoiperConfigurator/Releases'

function Get-LatestReleaseUrl {
    param(
        [string]$Owner,
        [string]$Repo,
        [string]$Branch,
        [string]$ReleasesPath,
        [string]$PreferredExt = '.ps1'
    )
    $headers = @{ 'User-Agent' = 'ZoiperUpdater'; Accept = 'application/vnd.github.v3+json' }
    try {
        $segments = $ReleasesPath.Split('/') | ForEach-Object { [Uri]::EscapeDataString($_) }
        $encodedPath = $segments -join '/'
        $uri = "https://api.github.com/repos/$Owner/$Repo/contents/$encodedPath"
        $items = Invoke-RestMethod -Uri $uri -Headers $headers -ErrorAction Stop
        $versionFolders = $items | Where-Object { $_.type -eq 'dir' -and $_.name -match '^[\d\.]+$' }
        $latestFolder = $versionFolders | ForEach-Object { [PSCustomObject]@{ Folder = $_.name; Version = [version]$_.name } } | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $latestFolder) { return $null }
        $folderUri = "https://api.github.com/repos/$Owner/$Repo/contents/$encodedPath/$($latestFolder.Folder)"
        $folderItems = Invoke-RestMethod -Uri $folderUri -Headers $headers -ErrorAction Stop
        $asset = $folderItems | Where-Object { $_.name -like "Zoiper Configurator*" -and $_.name -like "*${PreferredExt}" } | Select-Object -First 1
        if ($asset) {
            return "https://raw.githubusercontent.com/$Owner/$Repo/$Branch/$ReleasesPath/$($latestFolder.Folder)/$($asset.name)"
        }
    } catch {
        Write-Host "Failed to discover latest release: $_" -ForegroundColor Yellow
    }
    return $null
}


function Get-CurrentScriptPath {
    # Try $PSCommandPath first (works for .ps1 scripts)
    if ($PSCommandPath) { return $PSCommandPath }
    
    # Try $MyInvocation (works in some contexts)
    if ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) { 
        return $MyInvocation.MyCommand.Path 
    }
    
    # For ps2exe compiled scripts, get the executable path
    try {
        $proc = [System.Diagnostics.Process]::GetCurrentProcess()
        if ($proc.MainModule.FileName) {
            return $proc.MainModule.FileName
        }
    }
    catch { }
    
    return $null
}

function Get-VersionFromFile($path) {
    if (-not (Test-Path $path)) { return "0.0.0" }
    Start-Sleep -Milliseconds 200
    try {
        $text = [IO.File]::ReadAllText($path)
        if ($text -match 'ScriptVersion.*?(?<v>\d+\.\d+\.\d+)') { return $Matches.v }
        if ($text -match '(?<v>\d+\.\d+\.\d+)') { return $Matches.v }
        
        $gi = Get-Item $path
        if ($gi.VersionInfo.ProductVersion) { return $gi.VersionInfo.ProductVersion.Trim() }
        if ($gi.VersionInfo.FileVersion) { return $gi.VersionInfo.FileVersion.Trim() }
    }
    catch { }
    return "0.0.0"
}

function Invoke-SelfUpdate {
    param(
        [string]$UpdateUrl,
        [switch]$Force
    )

    if (-not $UpdateUrl) { Write-Verbose "No UpdateUrl configured."; return [PSCustomObject]@{ Status = 'NoUpdate'; RebootRequired = $false } }

    $proc = [System.Diagnostics.Process]::GetCurrentProcess()
    $currentProcessPath = $proc.MainModule.FileName
    $isExeRun = $currentProcessPath -and ($currentProcessPath.ToLower().EndsWith('.exe'))

    $targetPath = if ($isExeRun) { $currentProcessPath } else { Get-CurrentScriptPath }
    if (-not $targetPath) { Write-Warning "Cannot determine current script path. Self-update aborted."; return [PSCustomObject]@{ Status = 'NoUpdate'; RebootRequired = $false } }

    $ext = if ($UpdateUrl.ToLower().EndsWith('.exe') -or $isExeRun) { '.exe' } else { '.ps1' }
    $temp = Join-Path $env:TEMP ([IO.Path]::GetRandomFileName() + $ext)

    try {
        Invoke-WebRequest -Uri $UpdateUrl -OutFile $temp -UseBasicParsing -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to download update from $($UpdateUrl): $_"
        return [PSCustomObject]@{ Status = 'DownloadFailed'; RebootRequired = $false }
    }

    $localVersion = $ScriptVersion
    $remoteVersion = Get-VersionFromFile $temp

    Write-Host "Local Version:  $localVersion" -ForegroundColor Gray
    Write-Host "Remote Version: $remoteVersion" -ForegroundColor Gray

    $isNewer = $false
    if ($remoteVersion -and $localVersion) { $isNewer = [version]$remoteVersion -gt [version]$localVersion }

    if (-not $isNewer -and -not $Force) {
        try { Remove-Item -Path $temp -ErrorAction SilentlyContinue } catch { }
        return [PSCustomObject]@{ Status = 'NoUpdate'; RebootRequired = $false }
    }

    if ($Force -and -not $isNewer) { Write-Host "Forcing re-installation..." -ForegroundColor Yellow }
    else { Write-Host "Update found ($remoteVersion) — installing..." -ForegroundColor Cyan }

    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    $msg = "A new version ($remoteVersion) has been downloaded and is ready to be installed.`n`nPlease close the Zoiper Configurator and then click OK to complete installation. You will need to relaunch the application manually after installation."
    [System.Windows.Forms.MessageBox]::Show($msg, "Update Ready", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null

    $backup = "$targetPath.old"
    try {
        if (Test-Path -Path $targetPath) { Move-Item -Path $targetPath -Destination $backup -Force -ErrorAction Stop }
        Copy-Item -Path $temp -Destination $targetPath -Force -ErrorAction Stop
        if (Test-Path -Path $backup) { Remove-Item -Path $backup -Force -ErrorAction SilentlyContinue }
        return [PSCustomObject]@{ Status = 'Updated'; RebootRequired = $true }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("The update could not be installed automatically. Please close any running instances and copy the file:`n$temp`nto:`n$targetPath","Update Failed",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return [PSCustomObject]@{ Status = 'CopyFailed'; RebootRequired = $false }
    }
    finally {
        try { Remove-Item -Path $temp -ErrorAction SilentlyContinue } catch { }
    }
}

 
# Central update configuration and auto-check
$ScriptUpdateConfig = @{
    AutoCheck = $true
    UpdateUrl  = Get-LatestReleaseUrl -Owner $GitHubOwner -Repo $GitHubRepo -Branch $GitHubBranch -ReleasesPath $GitHubReleasesPath -PreferredExt '.ps1'
}

if ($ScriptUpdateConfig.AutoCheck -and $ScriptUpdateConfig.UpdateUrl) {
    $autoUpdateResult = Invoke-SelfUpdate -UpdateUrl $ScriptUpdateConfig.UpdateUrl
    if ($autoUpdateResult -and $autoUpdateResult.RebootRequired) { exit }
}

Write-Host "Starting Zoiper 5 Setup..." -ForegroundColor Cyan

# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Write-Host "Initializing UI components..." -ForegroundColor Cyan

# Create function to easily create styled labels
function New-StyledLabel ($text, $top, $isHeader = $false) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.AutoSize = $true
    if ($isHeader) {
        $label.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        $label.ForeColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
        $label.Location = New-Object System.Drawing.Point(20, $top)
    }
    else {
        $label.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $label.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
        $label.Location = New-Object System.Drawing.Point(22, $top)
    }
    return $label
}

# Create function to create modern styled buttons
function New-StyledButton ($text, $left, $top, $isPrimary = $true) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Location = New-Object System.Drawing.Point($left, $top)
    $btn.Size = New-Object System.Drawing.Size(100, 35)
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.FlatAppearance.BorderSize = 0
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    
    if ($isPrimary) {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215) # Modern Windows Blue
        $btn.ForeColor = [System.Drawing.Color]::White
        $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(0, 100, 190)
        $btn.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(0, 80, 160)
    }
    else {
        $btn.BackColor = [System.Drawing.Color]::White
        $btn.ForeColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
        $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
        $btn.FlatAppearance.BorderSize = 1
        $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    }
    
    return $btn
}

# Check for running Zoiper instance
if (Get-Process -Name "Zoiper5" -ErrorAction SilentlyContinue) {
    $closeForm = New-Object System.Windows.Forms.Form
    $closeForm.Text = "Action Required: Zoiper Running"
    $closeForm.Size = New-Object System.Drawing.Size(460, 200)
    $closeForm.StartPosition = "CenterScreen"
    $closeForm.FormBorderStyle = "FixedDialog"
    $closeForm.MaximizeBox = $false
    $closeForm.MinimizeBox = $true
    $closeForm.ShowInTaskbar = $true
    $closeForm.TopMost = $true
    $closeForm.BackColor = [System.Drawing.Color]::White

    $lblRunning = New-StyledLabel "Zoiper 5 is currently running." 20 $true
    $closeForm.Controls.Add($lblRunning)

    $lblInstruction = New-StyledLabel "Zoiper must be closed to proceed with the setup.`nClick 'Close Zoiper' to force close it and continue." 60
    $lblInstruction.Size = New-Object System.Drawing.Size(420, 50) 
    $closeForm.Controls.Add($lblInstruction)

    $btnCloseZoiper = New-StyledButton "Close Zoiper" 120 110 $true
    $btnCloseZoiper.Width = 100
    $btnCloseZoiper.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $closeForm.Controls.Add($btnCloseZoiper)
    $closeForm.AcceptButton = $btnCloseZoiper

    $btnExit = New-StyledButton "Exit" 240 110 $false
    $btnExit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $closeForm.Controls.Add($btnExit)
    $closeForm.CancelButton = $btnExit

    $closeResult = $closeForm.ShowDialog()

    if ($closeResult -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Force closing Zoiper..." -ForegroundColor Yellow
        Get-Process -Name "Zoiper5" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
    else {
        Write-Host "Setup cancelled by user." -ForegroundColor Red
        exit
    }
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Zoiper 5 Setup"
$form.Size = New-Object System.Drawing.Size(400, 360)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.ShowInTaskbar = $true
$form.BackColor = [System.Drawing.Color]::White

# Create TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(0, 0)
$tabControl.Size = New-Object System.Drawing.Size(384, 260)
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($tabControl)

# Tab 1: Configuration
$tabConfig = New-Object System.Windows.Forms.TabPage
$tabConfig.Text = "Configuration"
$tabConfig.BackColor = [System.Drawing.Color]::White
$tabControl.Controls.Add($tabConfig)

# Header
$header = New-StyledLabel "Zoiper 5 Configuration" 15 $true
$tabConfig.Controls.Add($header)

$subHeader = New-StyledLabel "Please enter your extension details below." 45
$subHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$tabConfig.Controls.Add($subHeader)

# Label: Extension Number
$labelExtension = New-StyledLabel "Extension Number" 85
$tabConfig.Controls.Add($labelExtension)

# TextBox: Extension Number
$textBoxExtension = New-Object System.Windows.Forms.TextBox
$textBoxExtension.Location = New-Object System.Drawing.Point(25, 110)
$textBoxExtension.Size = New-Object System.Drawing.Size(320, 26)
$textBoxExtension.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$tabConfig.Controls.Add($textBoxExtension)

# Label: Password
$labelPassword = New-StyledLabel "Password" 150
$tabConfig.Controls.Add($labelPassword)

# TextBox: Password
$textBoxPassword = New-Object System.Windows.Forms.TextBox
$textBoxPassword.Location = New-Object System.Drawing.Point(25, 175)
$textBoxPassword.Size = New-Object System.Drawing.Size(320, 26)
$textBoxPassword.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$textBoxPassword.PasswordChar = '*'
$tabConfig.Controls.Add($textBoxPassword)

# Tab 2: About
$tabAbout = New-Object System.Windows.Forms.TabPage
$tabAbout.Text = "About"
$tabAbout.BackColor = [System.Drawing.Color]::White
$tabControl.Controls.Add($tabAbout)

# About Content
$aboutHeader = New-StyledLabel "Zoiper Configurator" 20 $true
$tabAbout.Controls.Add($aboutHeader)

$versionLabel = New-StyledLabel "Version: $ScriptVersion" 50
$tabAbout.Controls.Add($versionLabel)

$vendorLabel = New-StyledLabel "Vendor: Skill On Net" 75
$tabAbout.Controls.Add($vendorLabel)

$descLabel = New-StyledLabel "This utility automates the configuration of Zoiper 5`nwith the Skill On Net SIP settings." 110
$descLabel.Size = New-Object System.Drawing.Size(330, 50)
$tabAbout.Controls.Add($descLabel)

# Button: Check for Updates in About Tab
$updateButton = New-StyledButton "Check for Updates" 25 170 $false
$updateButton.Size = New-Object System.Drawing.Size(150, 35)
$tabAbout.Controls.Add($updateButton)

$updateButton.Add_Click({
    if (-not $ScriptUpdateConfig.UpdateUrl) {
        [System.Windows.Forms.MessageBox]::Show("No update source configured.", "Update Check", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $updateResult = Invoke-SelfUpdate -UpdateUrl $ScriptUpdateConfig.UpdateUrl

    if ($updateResult -and $updateResult.Status -in @('DownloadFailed','CopyFailed')) {
        [System.Windows.Forms.MessageBox]::Show("The update could not be completed. Please try again later.", "Update Check", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    if ($updateResult -and $updateResult.RebootRequired) {
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Abort
        return
    }

    if (-not $updateResult -or $updateResult.Status -eq 'NoUpdate') {
        $msg = if (-not $updateResult) {
            "Could not find a newer version. Would you like to force reinstall the current version?"
        }
        else {
            "You are currently on the latest version. Would you like to force an update anyway?"
        }

        $res = [System.Windows.Forms.MessageBox]::Show($msg, "Update Check", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
            $forceResult = Invoke-SelfUpdate -UpdateUrl $ScriptUpdateConfig.UpdateUrl -Force
            if ($forceResult -and $forceResult.RebootRequired) {
                $form.DialogResult = [System.Windows.Forms.DialogResult]::Abort
            }
        }
    }
})

# Separator line
$line = New-Object System.Windows.Forms.Label
$line.Location = New-Object System.Drawing.Point(0, 265)
$line.Size = New-Object System.Drawing.Size(400, 1)
$line.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$form.Controls.Add($line)

# Button Panel Background
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(0, 266)
$buttonPanel.Size = New-Object System.Drawing.Size(400, 55)
$buttonPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$form.Controls.Add($buttonPanel)

# Button: OK (Save)
$okButton = New-StyledButton "Save" 170 10 $true
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$buttonPanel.Controls.Add($okButton)

# Button: Cancel
$cancelButton = New-StyledButton "Cancel" 280 10 $false
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$buttonPanel.Controls.Add($cancelButton)

$form.TopMost = $true

# Show the dialog

$result = $form.ShowDialog()


