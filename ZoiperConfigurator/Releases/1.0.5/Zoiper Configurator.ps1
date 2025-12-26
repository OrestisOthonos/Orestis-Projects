# Test
# Zoiper 5 Setup Helper
# This script prompts the user for Zoiper 5 credentials.
$ScriptVersion = '1.0.5'

# --- Self-update configuration ---
# Set `UpdateUrl` to your public release download URL or leave empty and use the
# authenticated update function below for private releases.
$ScriptUpdateConfig = @{ 
    # Example public releases URL:
    # 'https://github.com/OWNER/REPO/releases/latest/download/ZoiperConfigurator.exe'
    UpdateUrl    = '' 
    
    # Public GitHub Repo Configuration
    GitHubOwner  = 'OrestisOthonos' 
    GitHubRepo   = 'Orestis-Projects' 
    GitHubBranch = 'main'
    # Path to the releases folder in the repo
    GitHubPath   = 'ZoiperConfigurator/Releases'
    
    AutoCheck    = $true # Set to $true to check for updates automatically on start
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
        [switch]$RestartAfterUpdate,
        [switch]$Force
    )

    if (-not $UpdateUrl) { Write-Verbose "No UpdateUrl configured."; return }

    $proc = [System.Diagnostics.Process]::GetCurrentProcess()
    $currentProcessPath = $proc.MainModule.FileName
    $isExeRun = $currentProcessPath -and ($currentProcessPath.ToLower().EndsWith('.exe'))

    if ($isExeRun) { $targetPath = $currentProcessPath }
    else {
        $scriptPath = Get-CurrentScriptPath
        if (-not $scriptPath) { Write-Warning "Cannot determine current script path. Self-update aborted."; return }
        $targetPath = $scriptPath
    }

    $ext = '.ps1'
    if ($UpdateUrl.ToLower().EndsWith('.exe') -or $isExeRun) { $ext = '.exe' }
    $temp = Join-Path $env:TEMP ([IO.Path]::GetRandomFileName() + $ext)

    try {
        Invoke-WebRequest -Uri $UpdateUrl -OutFile $temp -UseBasicParsing -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to download update from $($UpdateUrl): $_"
        return
    }

    $localVersion = $ScriptVersion
    $remoteVersion = Get-VersionFromFile `"$temp`"

    Write-Host "Local Version:  $localVersion" -ForegroundColor Gray
    Write-Host "Remote Version: $remoteVersion" -ForegroundColor Gray

    $isNewer = $false
    if ($remoteVersion -and $localVersion) { $isNewer = [version]$remoteVersion -gt [version]$localVersion }

    if ($isNewer -or $Force) {
        if ($Force -and -not $isNewer) { Write-Host "Forcing re-installation..." -ForegroundColor Yellow }
        else { Write-Host "Update found ($remoteVersion) — installing..." -ForegroundColor Cyan }

        $updaterPath = Join-Path $env:TEMP ("zoiper_updater_" + [IO.Path]::GetRandomFileName() + ".ps1")
        $parentPid = $PID

        $updaterScript = @"
param(
    [string]
    `$Target,
    [string]
    `$Source,
    [int]
    `$ParentPid,
    [switch]
    `$Restart
)

while (Get-Process -Id `$ParentPid -ErrorAction SilentlyContinue) { Start-Sleep -Milliseconds 300 }

try { if (Test-Path -Path `$Target) { Remove-Item -Path `$Target -Force -ErrorAction Stop } } catch { }
Copy-Item -Path `$Source -Destination `$Target -Force
if (`$Restart) { 
    if (`$Target.ToLower().EndsWith('.ps1')) {
        Start-Process -FilePath 'powershell.exe' -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"`$Target`""
    } else {
        Start-Process -FilePath `$Target
    }
}
Remove-Item -Path `$Source -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 200
Remove-Item -Path `$MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue
"@

        $updaterScript | Out-File -FilePath $updaterPath -Encoding UTF8

        $processArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $updaterPath, '--', '-Target', "'$($targetPath.Replace("'", "''"))'", '-Source', "'$($temp.Replace("'", "''"))'", '-ParentPid', $parentPid)
        if ($RestartAfterUpdate) { $processArgs += '-Restart' }

        Start-Process -FilePath 'powershell' -ArgumentList $processArgs -WindowStyle Hidden

        Write-Host "Update downloaded; returning to trigger graceful exit." -ForegroundColor Yellow
        return [PSCustomObject]@{ Status = 'Updated'; RebootRequired = $true }
    }
    return [PSCustomObject]@{ Status = 'NoUpdate'; RebootRequired = $false }
}

# --- Helpers for private GitHub releases (PAT support) ---
function Get-GitHubToken {
    param(
        [string]$VaultEntryName = 'GitHubPAT'
    )

    if ($env:GITHUB_PAT) { return $env:GITHUB_PAT }

    if (Get-Command -Name Get-StoredCredential -ErrorAction SilentlyContinue) {
        try {
            $stored = Get-StoredCredential -Target $VaultEntryName -ErrorAction SilentlyContinue
            if ($stored -and $stored.Password) { return $stored.Password }
        }
        catch { }
    }

    return $null
}

# --- Public GitHub Download Functions ---

function Get-PublicRepoContent {
    param(
        [string]$Owner,
        [string]$Repo,
        [string]$Branch,
        [string]$Path,
        [string]$OutFile
    )
    
    # Use GitHub's raw content URL for public repos
    $uri = "https://raw.githubusercontent.com/$Owner/$Repo/$Branch/$Path"
    
    Write-Host "Downloading: $Path..." -ForegroundColor Gray
    Invoke-WebRequest -Uri $uri -OutFile $OutFile -UseBasicParsing -Method Get -ErrorAction Stop
}

function Get-LatestPublicGitHubPath {
    param(
        [string]$Owner,
        [string]$Repo,
        [string]$Branch,
        [string]$BasePath,
        [string]$PreferredExt = '.ps1'
    )

    $headers = @{ 'User-Agent' = 'ZoiperUpdater'; Accept = 'application/vnd.github.v3+json' }

    try {
        # List contents of the base path using GitHub API (works for public repos without auth)
        $segments = $BasePath.Split('/') | ForEach-Object { [Uri]::EscapeDataString($_) }
        $encodedPath = $segments -join '/'
        $uri = "https://api.github.com/repos/$Owner/$Repo/contents/$encodedPath"
        
        $items = Invoke-RestMethod -Uri $uri -Headers $headers -ErrorAction Stop
        
        # Filter for folders that look like versions (e.g., 1.0, 1.0.1)
        $versionFolders = $items | Where-Object { $_.type -eq 'dir' -and $_.name -match '^[\d\.]+$' }
        if (-not $versionFolders) { return $null }

        # Sort folders by version and pick the highest
        $latestFolder = $versionFolders | ForEach-Object { 
            try {
                [PSCustomObject]@{ Folder = $_; Version = [version]$_.name }
            }
            catch {
                # Skip folders that don't parse as valid versions
                Write-Verbose "Skipping invalid version folder: $($_.name)"
            }
        } | Where-Object { $_.Version } | Sort-Object Version -Descending | Select-Object -First 1

        if (-not $latestFolder) { return $null }

        # Now list the contents of that specific version folder to find our script/exe
        $folderUri = $latestFolder.Folder.url
        $folderItems = Invoke-RestMethod -Uri $folderUri -Headers $headers -ErrorAction Stop
        
        # Look for matching assets (prefer the preferred extension)
        $assets = $folderItems | Where-Object { $_.name -like "Zoiper Configurator.*" -or $_.name -like "ZoiperSetup.*" }
        $asset = ($assets | Where-Object { $_.name -like "*$PreferredExt" } | Select-Object -First 1)
        if (-not $asset) { $asset = $assets | Select-Object -First 1 }
        
        if ($asset) { return $asset.path }
    }
    catch {
        # If API fails (rate limit, etc.), try a smart fallback
        if ($_.Exception.Message -match '403|rate limit') {
            Write-Host "GitHub API rate limit reached. Trying direct version checks..." -ForegroundColor Yellow
            
            # First, try the current version (for force reinstalls)
            $currentVer = [version]$ScriptVersion
            $currentPath = "$BasePath/$ScriptVersion/Zoiper Configurator$PreferredExt"
            $currentUri = "https://raw.githubusercontent.com/$Owner/$Repo/$Branch/$currentPath"
            
            try {
                $null = Invoke-WebRequest -Uri $currentUri -UseBasicParsing -ErrorAction Stop -TimeoutSec 3
                Write-Host "Found current version $ScriptVersion (for reinstall)" -ForegroundColor Cyan
                return $currentPath
            }
            catch {
                # Current version not found, try next version
            }
            
            # Try the next incremental version
            $nextVer = "$($currentVer.Major).$($currentVer.Minor).$($currentVer.Build + 1)"
            $testPath = "$BasePath/$nextVer/Zoiper Configurator$PreferredExt"
            $testUri = "https://raw.githubusercontent.com/$Owner/$Repo/$Branch/$testPath"
            
            try {
                $null = Invoke-WebRequest -Uri $testUri -UseBasicParsing -ErrorAction Stop -TimeoutSec 3
                Write-Host "Found version $nextVer" -ForegroundColor Green
                return $testPath
            }
            catch {
                # Next version doesn't exist either
            }
        }
        Write-Warning "Discovery failed: $($_.Exception.Message)"
    }
    return $null
}

function Invoke-PublicUpdate {
    param(
        [string]$Owner,
        [string]$Repo,
        [string]$Branch,
        [string]$Path,
        [switch]$RestartAfterUpdate,
        [switch]$Force
    )

    # Determine if we're running as .ps1 or .exe by checking the actual script path
    $scriptPath = Get-CurrentScriptPath
    $isExeRun = $scriptPath -and ($scriptPath.ToLower().EndsWith('.exe'))

    if ($isExeRun) { 
        $targetPath = $scriptPath
    }
    else {
        if (-not $scriptPath) { Write-Warning "Cannot determine current script path. Update aborted."; return }
        $targetPath = $scriptPath
    }

    $ext = '.ps1'
    if ($targetPath.ToLower().EndsWith('.exe')) { $ext = '.exe' }
    $temp = Join-Path $env:TEMP ([IO.Path]::GetRandomFileName() + $ext)

    try {
        $discoveryPath = $Path
        # If Path points to the releases root, try to discover the latest versioned folder
        if ($Path -and $Path.EndsWith("Releases")) {
            Write-Host "Searching for latest version on GitHub..." -ForegroundColor Gray
            $discovered = Get-LatestPublicGitHubPath -Owner $Owner -Repo $Repo -Branch $Branch -BasePath $Path -PreferredExt $ext
            if ($discovered) { 
                $discoveryPath = $discovered 
                Write-Host "Discovered path: $discoveryPath" -ForegroundColor Cyan
            }
            else {
                Write-Warning "No version discovered, falling back to base path"
            }
        }

        if ($discoveryPath) {
            Write-Host "Attempting download from: $discoveryPath" -ForegroundColor Gray
            Get-PublicRepoContent -Owner $Owner -Repo $Repo -Branch $Branch -Path $discoveryPath -OutFile `"$temp`"
        }
        else {
            throw "No valid download path found"
        }
        Start-Sleep -Milliseconds 200
    }
    catch {
        Write-Warning "Failed to download update: $_"
        return
    }

    $localVersion = $ScriptVersion
    $remoteVersion = Get-VersionFromFile `"$temp`"

    Write-Host "Local Version:  $localVersion" -ForegroundColor Gray
    Write-Host "Remote Version: $remoteVersion" -ForegroundColor Gray

    $isNewer = $false
    if ($remoteVersion -and $localVersion) { $isNewer = [version]$remoteVersion -gt [version]$localVersion }

    if ($isNewer -or $Force) {
        if ($Force -and -not $isNewer) { Write-Host "Forcing re-installation..." -ForegroundColor Yellow }
        else { Write-Host "Update found ($remoteVersion) — installing..." -ForegroundColor Cyan }

        $updaterPath = Join-Path $env:TEMP ("zoiper_updater_" + [IO.Path]::GetRandomFileName() + ".ps1")
        $parentPid = $PID

        $updaterScript = @"
param(
    [string]
    `$Target,
    [string]
    `$Source,
    [int]
    `$ParentPid,
    [switch]
    `$Restart
)

while (Get-Process -Id `$ParentPid -ErrorAction SilentlyContinue) { Start-Sleep -Milliseconds 300 }

try { if (Test-Path -Path `$Target) { Remove-Item -Path `$Target -Force -ErrorAction Stop } } catch { }
Copy-Item -Path `$Source -Destination `$Target -Force
if (`$Restart) { 
    if (`$Target.ToLower().EndsWith('.ps1')) {
        Start-Process -FilePath 'powershell.exe' -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"`$Target`""
    } else {
        Start-Process -FilePath `$Target
    }
}
Remove-Item -Path `$Source -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 200
Remove-Item -Path `$MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue
"@

        $updaterScript | Out-File -FilePath $updaterPath -Encoding UTF8

        $processArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $updaterPath, '--', '-Target', "'$($targetPath.Replace("'", "''"))'", '-Source', "'$($temp.Replace("'", "''"))'", '-ParentPid', $parentPid)
        if ($RestartAfterUpdate) { $processArgs += '-Restart' }

        Start-Process -FilePath 'powershell' -ArgumentList $processArgs -WindowStyle Hidden

        Write-Host "Update downloaded; returning to trigger graceful exit." -ForegroundColor Yellow
        return [PSCustomObject]@{ Status = 'Updated'; RebootRequired = $true }
    }
    return [PSCustomObject]@{ Status = 'NoUpdate'; RebootRequired = $false }
}

# Auto-check if configured
if ($ScriptUpdateConfig.AutoCheck -eq $true) {
    Write-Host "Checking for updates..." -ForegroundColor Cyan
    if ($ScriptUpdateConfig.UpdateUrl) {
        Invoke-SelfUpdate -UpdateUrl $ScriptUpdateConfig.UpdateUrl
    }
    elseif ($ScriptUpdateConfig.GitHubOwner -and $ScriptUpdateConfig.GitHubRepo) {
        # Note: If an update happens here, script will exit. 
        # For auto-check, we usually want it to stay silent if failed.
        try {
            $result = Invoke-PublicUpdate -Owner $ScriptUpdateConfig.GitHubOwner -Repo $ScriptUpdateConfig.GitHubRepo -Branch $ScriptUpdateConfig.GitHubBranch -Path $ScriptUpdateConfig.GitHubPath -RestartAfterUpdate
            if ($result -and $result.RebootRequired) { exit }
        }
        catch {
            Write-Host "Update check skipped: $($_.Exception.Message)" -ForegroundColor Gray
        }
    }
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
$textBoxPassword.PasswordChar = "●"
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
        $updateResult = $null
        if ($ScriptUpdateConfig.UpdateUrl) {
            $updateResult = Invoke-SelfUpdate -UpdateUrl $ScriptUpdateConfig.UpdateUrl -RestartAfterUpdate
        }
        elseif ($ScriptUpdateConfig.GitHubOwner -and $ScriptUpdateConfig.GitHubRepo) {
            $updateResult = Invoke-PublicUpdate -Owner $ScriptUpdateConfig.GitHubOwner -Repo $ScriptUpdateConfig.GitHubRepo -Branch $ScriptUpdateConfig.GitHubBranch -Path $ScriptUpdateConfig.GitHubPath -RestartAfterUpdate
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("No update source configured.", "Update Check", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        if ($updateResult -and $updateResult.RebootRequired) {
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Abort
            return
        }

        # If no update was found OR download failed, offer force update
        if (-not $updateResult -or $updateResult.Status -eq 'NoUpdate') {
            $msg = if (-not $updateResult) {
                "Could not find a newer version. Would you like to force reinstall the current version?"
            }
            else {
                "You are currently on the latest version. Would you like to force an update anyway?"
            }
            
            $res = [System.Windows.Forms.MessageBox]::Show($msg, "Update Check", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
                $forceResult = $null
                if ($ScriptUpdateConfig.UpdateUrl) {
                    $forceResult = Invoke-SelfUpdate -UpdateUrl $ScriptUpdateConfig.UpdateUrl -RestartAfterUpdate -Force
                }
                else {
                    $forceResult = Invoke-PublicUpdate -Owner $ScriptUpdateConfig.GitHubOwner -Repo $ScriptUpdateConfig.GitHubRepo -Branch $ScriptUpdateConfig.GitHubBranch -Path $ScriptUpdateConfig.GitHubPath -RestartAfterUpdate -Force
                }

                if ($forceResult -and $forceResult.RebootRequired) {
                    $form.DialogResult = [System.Windows.Forms.DialogResult]::Abort
                    return
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

# Handle update-triggered exit
if ($result -eq [System.Windows.Forms.DialogResult]::Abort) {
    Write-Host "Update in progress. Exiting to allow file replacement..." -ForegroundColor Yellow
    exit
}

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $username = $textBoxExtension.Text
    $password = $textBoxPassword.Text

    Write-Host "Credentials captured successfully." -ForegroundColor Green
    Write-Host "Extension Number: $username" -ForegroundColor Yellow
    Write-Host "Password has been stored securely." -ForegroundColor Yellow
    
    # Define Config Path
    $configPath = "$env:APPDATA\Zoiper5\Config.xml"

    # Check if Config.xml exists, if not, try to generate it by running Zoiper
    if (-not (Test-Path $configPath)) {
        Write-Host "Config.xml not found. Zoiper 5 needs to run once to generate it." -ForegroundColor Cyan
        
        $zoiperPaths = @(
            "$env:ProgramFiles\Zoiper5\Zoiper5.exe",
            "${env:ProgramFiles(x86)}\Zoiper5\Zoiper5.exe"
        )

        $zoiperPdf = $null
        foreach ($path in $zoiperPaths) {
            if (Test-Path $path) {
                $zoiperPdf = $path
                break
            }
        }

        if ($zoiperPdf) {
            Write-Host "Starting Zoiper 5..." -ForegroundColor Cyan
            $process = Start-Process -FilePath $zoiperPdf -PassThru
            
            Write-Host "Waiting 10 seconds for Zoiper to initialize..." -ForegroundColor Cyan
            Start-Sleep -Seconds 10
            
            # --- Custom Instruction Dialog ---
            
            # Design Custom Form
            $instructForm = New-Object System.Windows.Forms.Form
            $instructForm.Text = "Action Required: Close Zoiper"
            $instructForm.Size = New-Object System.Drawing.Size(550, 220)
            $instructForm.StartPosition = "CenterScreen"
            $instructForm.FormBorderStyle = "FixedDialog"
            $instructForm.MaximizeBox = $false
            $instructForm.MinimizeBox = $true
            $instructForm.ShowInTaskbar = $true
            $instructForm.TopMost = $true # Always on top as requested
            $instructForm.BackColor = [System.Drawing.Color]::White

            # Instructions Label
            $lblVal = "Zoiper 5 has been launched to generate the configuration file.`n`n" +
            "IMPORTANT: Zoiper minimizes to the tray by default.`n`n" +
            "Please right-click the Zoiper icon in the system tray (near the clock) and select 'Exit'."
            
            $lbl = New-StyledLabel $lblVal 10
            $lbl.Size = New-Object System.Drawing.Size(510, 85)
            $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $instructForm.Controls.Add($lbl)

            # OK Button - Modernized
            $btnOK = New-StyledButton "I have closed Zoiper" 175 120 $true
            $btnOK.Size = New-Object System.Drawing.Size(200, 35) # Wider custom size
            $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $instructForm.Controls.Add($btnOK)
            $instructForm.AcceptButton = $btnOK

            # Show Dialog
            $instructForm.ShowDialog() | Out-Null

            Write-Host "Waiting for you to close Zoiper 5 via the system tray..." -ForegroundColor Yellow
            $process.WaitForExit()
            
            # Wait a moment for file release
            Start-Sleep -Seconds 2
            Write-Host "Zoiper closed. Proceeding with configuration..." -ForegroundColor Green
        }
        else {
            Write-Warning "Could not find Zoiper 5 executable. Please start Zoiper manually once to generate the configuration."
        }
    }

    if (Test-Path $configPath) {
        try {
            # Load the XML
            [xml]$xml = Get-Content $configPath

            # Check for existing accounts
            $existingAccounts = $xml.SelectNodes("//options/accounts/account")
            if ($existingAccounts.Count -gt 0) {
                # Create Conflict Dialog
                $conflictForm = New-Object System.Windows.Forms.Form
                $conflictForm.Text = "Account Conflict"
                $conflictForm.Size = New-Object System.Drawing.Size(460, 220)
                $conflictForm.StartPosition = "CenterScreen"
                $conflictForm.FormBorderStyle = "FixedDialog"
                $conflictForm.MaximizeBox = $false
                $conflictForm.MinimizeBox = $true
                $conflictForm.ShowInTaskbar = $true
                $conflictForm.TopMost = $true
                $conflictForm.BackColor = [System.Drawing.Color]::White

                $lblConflict = New-StyledLabel "Existing Account Found" 20 $true
                $lblConflict.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0) # Dark Red for warning
                $conflictForm.Controls.Add($lblConflict)

                $lblConflictMsg = New-StyledLabel "Configuration file already contains an account.`nDo you want to replace it with the new one?" 60
                $lblConflictMsg.Size = New-Object System.Drawing.Size(420, 50)
                $conflictForm.Controls.Add($lblConflictMsg)

                # Buttons
                $btnReplace = New-StyledButton "Replace" 120 120 $true
                $btnReplace.BackColor = [System.Drawing.Color]::FromArgb(200, 0, 0) # Red for destructive action
                $btnReplace.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
                $btnReplace.DialogResult = [System.Windows.Forms.DialogResult]::Yes
                $conflictForm.Controls.Add($btnReplace)
                $conflictForm.AcceptButton = $btnReplace

                $btnConflictExit = New-StyledButton "Exit" 240 120 $false
                $btnConflictExit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $conflictForm.Controls.Add($btnConflictExit)
                $conflictForm.CancelButton = $btnConflictExit

                $conflictResult = $conflictForm.ShowDialog()

                if ($conflictResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Write-Host "Removing existing accounts..." -ForegroundColor Yellow
                    $accountsContainer = $xml.SelectSingleNode("//options/accounts")
                    if ($accountsContainer) {
                        $accountsContainer.RemoveAll()
                    }
                }
                else {
                    Write-Host "Setup cancelled by user (Existing account retained)." -ForegroundColor Red
                    exit
                }
            }

            # Define the specific domain
            $domain = "sipout.kpaxmarketing.com"
            $accountName = "$username@$domain"

            # Create the account node structure
            # Updated based on user feedback to include correct domain and account name
            $newAccountXml = @"
      <account>
        <ident>$(New-Guid)</ident>
        <name>$accountName</name>
        <save_username>true</save_username>
        <username>$username</username>
        <save_password>true</save_password>
        <password>$password</password>
        <register_on_startup>true</register_on_startup>
        <active>true</active>
        <protocol>sip</protocol>
        <SIP_domain>$domain</SIP_domain>
        <SIP_transport_type>udp</SIP_transport_type>
        <stun>
          <use_stun>disabled</use_stun>
          <stun_host>stun.zoiper.com</stun_host>
          <stun_port>3478</stun_port>
          <stun_refresh_period>30</stun_refresh_period>
        </stun>
      </account>
"@
            # Import the new node
            $accountNode = $xml.ImportNode(([xml]$newAccountXml).DocumentElement, $true)

            # Check if <accounts> node exists, create if not (though it should exist)
            $accountsNode = $xml.SelectSingleNode("//options/accounts")
            if ($null -eq $accountsNode) {
                Write-Host "Accounts node not found, creating..." -ForegroundColor Yellow
                $accountsNode = $xml.CreateElement("accounts")
                $xml.options.AppendChild($accountsNode) | Out-Null
            }

            # Append the new account
            $accountsNode.AppendChild($accountNode) | Out-Null

            # --- Apply Additional Settings ---

            # General Settings
            $minToTray = $xml.SelectSingleNode("//options/general/minimize_to_tray")
            if ($minToTray) { $minToTray.InnerText = "false" }

            $minOnClose = $xml.SelectSingleNode("//options/general/minimize_on_close")
            if ($minOnClose) { $minOnClose.InnerText = "true" }

            $newCallAutoPopup = $xml.SelectSingleNode("//options/general/new_call_auto_popup")
            if ($newCallAutoPopup) { $newCallAutoPopup.InnerText = "true" }

            # GUI Settings
            $upgradeWizard = $xml.SelectSingleNode("//options/gui/show_upgrade_wizard")
            if ($upgradeWizard) { $upgradeWizard.InnerText = "false" }

            # STUN Settings - Disable STUN
            $useStun = $xml.SelectSingleNode("//options/stun/use_stun")
            if ($useStun) { $useStun.InnerText = "disabled" }

            # Chat Settings
            $newMessageBlink = $xml.SelectSingleNode("//options/chat/new_message_blink")
            if ($newMessageBlink) { $newMessageBlink.InnerText = "false" }

            # --- End Additional Settings ---

            # Save the file
            $xml.Save($configPath)
            Write-Host "Example configuration updated successfully!" -ForegroundColor Green
            Write-Host "New account added to $configPath" -ForegroundColor Gray

            # Send Notification to Teams via Power Automate Webhook
            try {
                $webhookUrl = "https://default56c1058d26dd4299b9d2d25d76787d.5d.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/dab2781e2ebe4217923387dadef86460/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=r2G91WZ2yvQrAH7DWg1FOHooq3rjiVXwUTZdzUrD05c"
                
                $payload = @{
                    windows_user     = $env:USERNAME
                    computer_name    = $env:COMPUTERNAME
                    zoiper_extension = $username
                } | ConvertTo-Json

                Invoke-RestMethod -Uri $webhookUrl -Method Post -Body $payload -ContentType "application/json" | Out-Null
                Write-Host "Notification sent to Teams successfully." -ForegroundColor Green
                
                # Notify User of Success (Always on Top)
                # Create Custom Modern Success Dialog
                $successForm = New-Object System.Windows.Forms.Form
                $successForm.Text = "Setup Complete"
                $successForm.Size = New-Object System.Drawing.Size(400, 200)
                $successForm.StartPosition = "CenterScreen"
                $successForm.FormBorderStyle = "FixedDialog"
                $successForm.MaximizeBox = $false
                $successForm.MinimizeBox = $true
                $successForm.ShowInTaskbar = $true
                $successForm.TopMost = $true
                $successForm.BackColor = [System.Drawing.Color]::White

                # Success Icon (using label with specific color/font)
                $lblSuccessHeader = New-StyledLabel "Success!" 20 $true
                $lblSuccessHeader.ForeColor = [System.Drawing.Color]::SeaGreen
                $successForm.Controls.Add($lblSuccessHeader)

                # Success Message
                $lblSuccessMsg = New-StyledLabel "Zoiper 5 setup has been completed successfully!" 60
                $lblSuccessMsg.Size = New-Object System.Drawing.Size(360, 50)
                $successForm.Controls.Add($lblSuccessMsg)

                # Buttons
                $btnStartZoiper = New-StyledButton "Start Zoiper" 70 110 $true
                $btnStartZoiper.Width = 120
                $btnStartZoiper.DialogResult = [System.Windows.Forms.DialogResult]::Yes
                $successForm.Controls.Add($btnStartZoiper)
                $successForm.AcceptButton = $btnStartZoiper

                $btnClose = New-StyledButton "Close Setup" 210 110 $false
                $btnClose.Width = 120
                $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::No
                $successForm.Controls.Add($btnClose)
                $successForm.CancelButton = $btnClose

                $finalResult = $successForm.ShowDialog()

                if ($finalResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Write-Host "Starting Zoiper 5..." -ForegroundColor Cyan
                    $zoiperPaths = @(
                        "$env:ProgramFiles\Zoiper5\Zoiper5.exe",
                        "${env:ProgramFiles(x86)}\Zoiper5\Zoiper5.exe"
                    )
                    foreach ($path in $zoiperPaths) {
                        if (Test-Path $path) {
                            Start-Process -FilePath $path
                            break
                        }
                    }
                }
            }
            catch {
                Write-Warning "Failed to send notification to Teams: $_"
            }

        }
        catch {
            Write-Error "Failed to update Config.xml: $_"
        }
    }
    else {
        Write-Warning "Config.xml not found at $configPath. Skipping configuration update."
    }
}
else {
    Write-Host "Failed to capture credentials. Please try again." -ForegroundColor Red
}
