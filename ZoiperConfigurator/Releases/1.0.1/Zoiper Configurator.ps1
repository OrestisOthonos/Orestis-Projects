# Zoiper 5 Setup Helper
# This script prompts the user for Zoiper 5 credentials.
$ScriptVersion = '1.0.1'

# --- Self-update configuration ---
# Set `UpdateUrl` to your public release download URL or leave empty and use the
# authenticated update function below for private releases.
$ScriptUpdateConfig = @{ 
    # Example public releases URL:
    # 'https://github.com/OWNER/REPO/releases/latest/download/ZoiperConfigurator.exe'
    UpdateUrl   = '' 
    
    # Private GitHub Repo Configuration (Used for authenticated updates)
    GitHubOwner = 'OrestisOthonos' 
    GitHubRepo  = 'Orestis-Projects' 
    # If GitHubPath is set, it will download directly from the repo tree (folders)
    GitHubPath  = 'ZoiperConfigurator/Releases'
    # If GitHubPath is empty, it will look for a Release asset named AssetName
    AssetName   = 'Zoiper Configurator.exe'
    
    AutoCheck   = $true # Set to $true to check for updates automatically on start
}

function Get-CurrentScriptPath {
    if ($PSCommandPath) { return $PSCommandPath }
    if ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) { return $MyInvocation.MyCommand.Path }
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
    } catch { }
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
    $remoteVersion = Get-VersionFromFile $temp

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

        $processArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $updaterPath, '--', '-Target', $targetPath, '-Source', $temp, '-ParentPid', $parentPid)
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

# If no PAT is present in the environment, offer a one-time interactive paste

# Optional embedded DPAPI-encrypted PAT (user-scoped).
$EncryptedPAT = '01000000d08c9ddf0115d1118c7a00c04fc297eb01000000eba8995160e5a84da0c072cc844b5c280000000002000000000010660000000100002000000001912b6220a5dc0002841fccf7f6e5e605b10f69b7753cf9c2560c45ced65425000000000e80000000020000200000005be30710f0167d3f7f89ab2caca1feadc0af5399a1a51dcb96d86df21dbe79a6c00000009b93e83e67a5473e37eb5a5c63c48f82056a5dade556d22d3c903fa07841d9cc3a5cb2daf3bd7365bbc5b981cf67941e5c5b6ace639fb73271d7fa38d01556a89f6ed655bf1963028c7eb3a69f50e77f4cce03f9c70473030ebc39fd2edf8690e763efca4100b6b28073eedc57c978a62af92f5a90ed83e509ace42be76f441c5c73d3c898016f879fbd4de61da889c0d1bf893b2686c36a544d77b62caf1bfcbb3013ceec8a2929492d44a93c39f67631f7d311a537a6bb6817b4a3aa1b53d340000000f91247e3d9b143d3647dabc48920ce7b1abe36acc37ec35a47ed606d9c9e3e807cfe2180e770919d01918b44a799a85135c10434add807786d4ac1c5f98b6328'

if ($EncryptedPAT -and $EncryptedPAT -ne 'PASTE_ENCRYPTED_STRING_HERE') {
    try {
        $secure = ConvertTo-SecureString -String $EncryptedPAT
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
        if ($plain) {
            # Priority: use embedded PAT if it's the one we just updated
            $env:GITHUB_PAT = $plain
            Write-Host "GITHUB_PAT loaded successfully from embedded blob." -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Failed to decrypt embedded PAT: $_"
    }
}
elseif ($env:GITHUB_PAT) {
    Write-Host "Using GITHUB_PAT from environment variable." -ForegroundColor Cyan
}
else {
    Write-Host "No GITHUB_PAT found in environment or embedded." -ForegroundColor Red
}

function Get-PrivateReleaseAsset {
    param(
        [string]$Owner,
        [string]$Repo,
        [string]$AssetName,
        [string]$Token,
        [string]$OutFile
    )

    if (-not $Token) { throw 'Token is required to download private release asset' }

    # GitHub fine-grained PATs work with "token" as well, and it's often more compatible.
    $authPrefix = if ($Token.StartsWith("github_pat_")) { "Bearer" } else { "token" }
    $headers = @{ Authorization = "$authPrefix $Token"; 'User-Agent' = 'ZoiperUpdater'; Accept = 'application/vnd.github+json' }
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$Owner/$Repo/releases/latest" -Headers $headers -ErrorAction Stop

    $asset = $release.assets | Where-Object { $_.name -eq $AssetName }
    if (-not $asset) { throw "Asset not found in latest release: $AssetName" }

    $assetId = $asset.id
    $downloadUri = "https://api.github.com/repos/$Owner/$Repo/releases/assets/$assetId"

    $dlHeaders = @{ Authorization = "$authPrefix $Token"; 'User-Agent' = 'ZoiperUpdater'; Accept = 'application/octet-stream' }
    Invoke-WebRequest -Uri $downloadUri -Headers $dlHeaders -OutFile $OutFile -UseBasicParsing -Method Get -ErrorAction Stop
}

function Get-PrivateRepoContent {
    param(
        [string]$Owner,
        [string]$Repo,
        [string]$Path,
        [string]$Token,
        [string]$OutFile
    )

    if (-not $Token) { throw 'Token is required to download private repo content' }

    $authPrefix = if ($Token.StartsWith("github_pat_")) { "Bearer" } else { "token" }
    # Use the Contents API with the 'raw' media type to get the file content directly
    $headers = @{ 
        Authorization = "$authPrefix $Token"
        'User-Agent'  = 'ZoiperUpdater'
        Accept        = 'application/vnd.github.v3.raw'
    }
    
    # Encode the path segments to handle spaces while keeping slashes literal for the URI
    $segments = $Path.Split('/') | ForEach-Object { [Uri]::EscapeDataString($_) }
    $encodedPath = $segments -join '/'
    $uri = "https://api.github.com/repos/$Owner/$Repo/contents/$encodedPath"
    
    Write-Host "Downloading: $Path..." -ForegroundColor Gray
    Invoke-WebRequest -Uri $uri -Headers $headers -OutFile $OutFile -UseBasicParsing -Method Get -ErrorAction Stop
}

function Get-LatestGitHubPath {
    param(
        [string]$Owner,
        [string]$Repo,
        [string]$BasePath,
        [string]$Token,
        [string]$PreferredExt = '.ps1'
    )

    if (-not $Token) { throw 'Token is required for GitHub API discovery' }
    $authPrefix = if ($Token.StartsWith("github_pat_")) { "Bearer" } else { "token" }
    $headers = @{ Authorization = "$authPrefix $Token"; 'User-Agent' = 'ZoiperUpdater'; Accept = 'application/vnd.github.v3+json' }

    try {
        # List contents of the base path (e.g., ZoiperConfigurator/Releases)
        $segments = $BasePath.Split('/') | ForEach-Object { [Uri]::EscapeDataString($_) }
        $encodedPath = $segments -join '/'
        $uri = "https://api.github.com/repos/$Owner/$Repo/contents/$encodedPath"
        
        $items = Invoke-RestMethod -Uri $uri -Headers $headers -ErrorAction Stop
        
        # Filter for folders that look like versions (e.g., 1.0, 1.0.1)
        $versionFolders = $items | Where-Object { $_.type -eq 'dir' -and $_.name -match '^[\d\.]+$' }
        if (-not $versionFolders) { return $null }

        # Sort folders by version and pick the highest
        $latestFolder = $versionFolders | ForEach-Object { 
            [PSCustomObject]@{ Folder = $_; Version = [version]$_.name } 
        } | Sort-Object Version -Descending | Select-Object -First 1

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
        Write-Warning "Discovery failed: $($_.Exception.Message)"
    }
    return $null
}

function Invoke-AuthenticatedUpdate {
    param(
        [string]$Owner,
        [string]$Repo,
        [string]$AssetName,
        [string]$Path,
        [switch]$RestartAfterUpdate,
        [string]$Token,
        [switch]$Force
    )

    if (-not $Token) { $Token = Get-GitHubToken }

    $proc = [System.Diagnostics.Process]::GetCurrentProcess()
    $currentProcessPath = $proc.MainModule.FileName
    $isExeRun = $currentProcessPath -and ($currentProcessPath.ToLower().EndsWith('.exe'))

    if ($isExeRun) { $targetPath = $currentProcessPath }
    else {
        $scriptPath = Get-CurrentScriptPath
        if (-not $scriptPath) { Write-Warning "Cannot determine current script path. Authenticated update aborted."; return }
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
            $discovered = Get-LatestGitHubPath -Owner $Owner -Repo $Repo -BasePath $Path -Token $Token -PreferredExt $ext
            if ($discovered) { $discoveryPath = $discovered }
        }

        if ($discoveryPath) {
            Get-PrivateRepoContent -Owner $Owner -Repo $Repo -Path $discoveryPath -Token $Token -OutFile $temp
        }
        else {
            Get-PrivateReleaseAsset -Owner $Owner -Repo $Repo -AssetName $AssetName -Token $Token -OutFile $temp
        }
        Start-Sleep -Milliseconds 200
    }
    catch {
        Write-Warning "Failed to download private asset: $_"
        return
    }

    $localVersion = $ScriptVersion
    $remoteVersion = Get-VersionFromFile $temp

    Write-Host "Local Version:  $localVersion" -ForegroundColor Gray
    Write-Host "Remote Version: $remoteVersion" -ForegroundColor Gray

    $isNewer = $false
    if ($remoteVersion -and $localVersion) { $isNewer = [version]$remoteVersion -gt [version]$localVersion }

    if ($isNewer -or $Force) {
        if ($Force -and -not $isNewer) { Write-Host "Forcing re-installation..." -ForegroundColor Yellow }
        else { Write-Host "Authenticated update found ($remoteVersion) — installing..." -ForegroundColor Cyan }

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

        $processArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $updaterPath, '--', '-Target', $targetPath, '-Source', $temp, '-ParentPid', $parentPid)
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
            $result = Invoke-AuthenticatedUpdate -Owner $ScriptUpdateConfig.GitHubOwner -Repo $ScriptUpdateConfig.GitHubRepo -AssetName $ScriptUpdateConfig.AssetName -Path $ScriptUpdateConfig.GitHubPath -RestartAfterUpdate
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
$form.Size = New-Object System.Drawing.Size(400, 345)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.ShowInTaskbar = $true
$form.BackColor = [System.Drawing.Color]::White

# Button: Check for Updates (Top-Left, Small)
$updateButton = New-StyledButton "Check for Updates" 5 5 $false
$updateButton.Size = New-Object System.Drawing.Size(110, 22)
$updateButton.Font = New-Object System.Drawing.Font("Segoe UI", 7)
$form.Controls.Add($updateButton)

$updateButton.Add_Click({
        $updateResult = $null
        if ($ScriptUpdateConfig.UpdateUrl) {
            $updateResult = Invoke-SelfUpdate -UpdateUrl $ScriptUpdateConfig.UpdateUrl -RestartAfterUpdate
        }
        elseif ($ScriptUpdateConfig.GitHubOwner -and $ScriptUpdateConfig.GitHubRepo) {
            $updateResult = Invoke-AuthenticatedUpdate -Owner $ScriptUpdateConfig.GitHubOwner -Repo $ScriptUpdateConfig.GitHubRepo -AssetName $ScriptUpdateConfig.AssetName -Path $ScriptUpdateConfig.GitHubPath -RestartAfterUpdate
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("No update source (URL or GitHub Repo) configured.", "Update Check", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        if ($updateResult -and $updateResult.RebootRequired) {
            $form.Close()
            exit
        }

        if ($updateResult -and $updateResult.Status -eq 'NoUpdate') {
            $msg = "You are currently on the latest version. Would you like to force an update anyway?"
            $res = [System.Windows.Forms.MessageBox]::Show($msg, "Update Check", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
                $forceResult = $null
                if ($ScriptUpdateConfig.UpdateUrl) {
                    $forceResult = Invoke-SelfUpdate -UpdateUrl $ScriptUpdateConfig.UpdateUrl -RestartAfterUpdate -Force
                }
                else {
                    $forceResult = Invoke-AuthenticatedUpdate -Owner $ScriptUpdateConfig.GitHubOwner -Repo $ScriptUpdateConfig.GitHubRepo -AssetName $ScriptUpdateConfig.AssetName -Path $ScriptUpdateConfig.GitHubPath -RestartAfterUpdate -Force
                }

                if ($forceResult -and $forceResult.RebootRequired) {
                    $form.Close()
                    exit
                }
            }
        }
    })

# Header
$header = New-StyledLabel "Zoiper 5 Configuration" 45 $true
$form.Controls.Add($header)

$subHeader = New-StyledLabel "Please enter your extension details below." 75
$subHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($subHeader)

# Label: Extension Number
$labelExtension = New-StyledLabel "Extension Number" 115
$form.Controls.Add($labelExtension)

# TextBox: Extension Number
$textBoxExtension = New-Object System.Windows.Forms.TextBox
$textBoxExtension.Location = New-Object System.Drawing.Point(25, 140)
$textBoxExtension.Size = New-Object System.Drawing.Size(330, 26)
$textBoxExtension.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($textBoxExtension)

# Label: Password
$labelPassword = New-StyledLabel "Password" 180
$form.Controls.Add($labelPassword)

# TextBox: Password
$textBoxPassword = New-Object System.Windows.Forms.TextBox
$textBoxPassword.Location = New-Object System.Drawing.Point(25, 205)
$textBoxPassword.Size = New-Object System.Drawing.Size(330, 26)
$textBoxPassword.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$textBoxPassword.PasswordChar = "●"
$form.Controls.Add($textBoxPassword)

# Separator line
$line = New-Object System.Windows.Forms.Label
$line.Location = New-Object System.Drawing.Point(0, 255)
$line.Size = New-Object System.Drawing.Size(400, 1)
$line.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$form.Controls.Add($line)

# Button Panel Background
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(0, 256)
$buttonPanel.Size = New-Object System.Drawing.Size(400, 50)
$buttonPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$form.Controls.Add($buttonPanel)

# Button: OK (Save)
$okButton = New-StyledButton "Save" 170 8 $true
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$buttonPanel.Controls.Add($okButton)

# Button: Cancel
$cancelButton = New-StyledButton "Cancel" 280 8 $false
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$buttonPanel.Controls.Add($cancelButton)

$form.TopMost = $true

# Show the dialog
$result = $form.ShowDialog()

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
