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

function Invoke-BatchUpdater {
    param(
        [string]$UpdateUrl,
        [version]$RemoteVersion,
        [string]$TargetPath,
        [int]$ParentPid
    )

    if (-not $UpdateUrl -or -not $TargetPath) { return }

    $batchPath = Join-Path $env:TEMP ("zoiper_update_" + [IO.Path]::GetRandomFileName() + ".bat")
    $downloadPath = Join-Path $env:TEMP ("zoiper_new_" + [IO.Path]::GetRandomFileName() + ".ps1")

    $batchContent = @"
@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -Uri `"$UpdateUrl`" -OutFile `"$downloadPath`" -UseBasicParsing -ErrorAction Stop; Wait-Process -Id $ParentPid -ErrorAction SilentlyContinue; Move-Item -Path `"$downloadPath`" -Destination `"$TargetPath`" -Force; Start-Process -FilePath `"powershell.exe`" -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File',`"$TargetPath`"; } catch { exit 1 }"
endlocal
"@

    Set-Content -Path $batchPath -Value $batchContent -Encoding ASCII
    Start-Process -FilePath $batchPath -WindowStyle Hidden
}

function Invoke-UpdateCheck {
    param(
        [switch]$SilentIfLatest,
        [switch]$Force
    )

    $latestUrl = Get-LatestReleaseUrl -Owner $GitHubOwner -Repo $GitHubRepo -Branch $GitHubBranch -ReleasesPath $GitHubReleasesPath -PreferredExt '.ps1'
    if (-not $latestUrl) {
        if (-not $SilentIfLatest) {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
            [System.Windows.Forms.MessageBox]::Show("Could not determine the latest release.", "Update Check", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        return 'NoUpdate'
    }

    $checkTemp = Join-Path $env:TEMP ("zoiper_ver_" + [IO.Path]::GetRandomFileName() + ".ps1")
    $remoteVersion = $null
    try {
        Invoke-WebRequest -Uri $latestUrl -OutFile $checkTemp -UseBasicParsing -ErrorAction Stop
        $remoteVersion = Get-VersionFromFile $checkTemp
    }
    catch {
        if (-not $SilentIfLatest) {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
            [System.Windows.Forms.MessageBox]::Show("Could not download the latest version.", "Update Check", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        }
        return 'NoUpdate'
    }
    finally { try { Remove-Item -Path $checkTemp -ErrorAction SilentlyContinue } catch { } }

    if (-not $Force) {
        if (-not $remoteVersion -or ([version]$remoteVersion -le [version]$ScriptVersion)) {
            if (-not $SilentIfLatest) {
                Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
                [System.Windows.Forms.MessageBox]::Show("You are already on the latest version.", "Update Check", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            }
            return 'NoUpdate'
        }
    }

    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    $targetPath = Get-CurrentScriptPath
    $targetPath = Get-CurrentScriptPath
    if (-not $targetPath) {
        if (-not $SilentIfLatest) {
            [System.Windows.Forms.MessageBox]::Show("Cannot determine the current script path. Update aborted.", "Update Check", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
        return 'NoUpdate'
    }

    $dialogText = if ($Force -and $remoteVersion) {
        "Forcing reinstall of version $remoteVersion. Zoiper Configurator will close to update now."
    } elseif ($remoteVersion) {
        "A new version ($remoteVersion) is available. Zoiper Configurator will close to update now."
    } else {
        "An update will be applied. Zoiper Configurator will close to update now."
    }
    [System.Windows.Forms.MessageBox]::Show($dialogText, "Update Available", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    Invoke-BatchUpdater -UpdateUrl $latestUrl -RemoteVersion $remoteVersion -TargetPath $targetPath -ParentPid $PID
    return 'Updating'
}

# Startup update check: compare versions and offload replacement to a batch file
$startupUpdate = Invoke-UpdateCheck -SilentIfLatest
if ($startupUpdate -eq 'Updating') { exit }

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
    $updateStatus = Invoke-UpdateCheck
    if ($updateStatus -eq 'Updating') { 
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Abort 
        return
    }

    # Offer force reinstall when no newer version is found
    if ($updateStatus -eq 'NoUpdate') {
        $res = [System.Windows.Forms.MessageBox]::Show("No newer version found. Force reinstall the latest release anyway?", "Update Check", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
            $forceStatus = Invoke-UpdateCheck -Force
            if ($forceStatus -eq 'Updating') { $form.DialogResult = [System.Windows.Forms.DialogResult]::Abort }
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


