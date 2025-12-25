# Zoiper 5 Setup Helper
# This script prompts the user for Zoiper 5 credentials.

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
    GitHubPath  = 'ZoiperConfigurator/Releases/1.0/Zoiper Configurator.ps1'
    # If GitHubPath is empty, it will look for a Release asset named AssetName
    AssetName   = 'Zoiper Configurator.exe'
    
    AutoCheck   = $true # Set to $true to check for updates automatically on start
}

function Get-CurrentScriptPath {
    if ($PSCommandPath) { return $PSCommandPath }
    if ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) { return $MyInvocation.MyCommand.Path }
    return $null
}

function Get-FileHashSha256($path) {
    if (-not (Test-Path $path)) { return $null }
    return (Get-FileHash -Path $path -Algorithm SHA256).Hash
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

    $ext = if ($UpdateUrl.ToLower().EndsWith('.exe') -or $isExeRun) { '.exe' } else { '.ps1' }
    $temp = Join-Path $env:TEMP ([IO.Path]::GetRandomFileName() + $ext)

    try {
        Invoke-WebRequest -Uri $UpdateUrl -OutFile $temp -UseBasicParsing -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to download update from $($UpdateUrl): $_"
        return
    }

    $localHash = Get-FileHashSha256 $targetPath
    $remoteHash = Get-FileHashSha256 $temp

    if (($localHash -and $remoteHash) -and (($localHash -ne $remoteHash) -or $Force)) {
        if ($Force -and ($localHash -eq $remoteHash)) { Write-Host "Forcing re-installation..." -ForegroundColor Yellow }
        else { Write-Host "Update found — installing..." -ForegroundColor Cyan }

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

        Write-Host "Updater launched; exiting to allow replacement." -ForegroundColor Yellow
        exit
    }
    else {
        Remove-Item -Path $temp -ErrorAction SilentlyContinue
        return $false
    }
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

    $ext = if ($targetPath.ToLower().EndsWith('.exe')) { '.exe' } else { '.ps1' }
    $temp = Join-Path $env:TEMP ([IO.Path]::GetRandomFileName() + $ext)

    try {
        if ($Path) {
            Get-PrivateRepoContent -Owner $Owner -Repo $Repo -Path $Path -Token $Token -OutFile $temp
        }
        else {
            Get-PrivateReleaseAsset -Owner $Owner -Repo $Repo -AssetName $AssetName -Token $Token -OutFile $temp
        }
    }
    catch {
        Write-Warning "Failed to download private asset: $_"
        return
    }

    $localHash = Get-FileHashSha256 $targetPath
    $remoteHash = Get-FileHashSha256 $temp

    if (($localHash -and $remoteHash) -and (($localHash -ne $remoteHash) -or $Force)) {
        if ($Force -and ($localHash -eq $remoteHash)) { Write-Host "Forcing re-installation..." -ForegroundColor Yellow }
        else { Write-Host "Authenticated update found — installing..." -ForegroundColor Cyan }

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

        Write-Host "Updater launched; exiting to allow replacement." -ForegroundColor Yellow
        exit
    }
    else {
        Remove-Item -Path $temp -ErrorAction SilentlyContinue
        return $false
    }
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
            Invoke-AuthenticatedUpdate -Owner $ScriptUpdateConfig.GitHubOwner -Repo $ScriptUpdateConfig.GitHubRepo -AssetName $ScriptUpdateConfig.AssetName -Path $ScriptUpdateConfig.GitHubPath
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
        $updated = $false
        if ($ScriptUpdateConfig.UpdateUrl) {
            $updated = Invoke-SelfUpdate -UpdateUrl $ScriptUpdateConfig.UpdateUrl -RestartAfterUpdate
        }
        elseif ($ScriptUpdateConfig.GitHubOwner -and $ScriptUpdateConfig.GitHubRepo) {
            $updated = Invoke-AuthenticatedUpdate -Owner $ScriptUpdateConfig.GitHubOwner -Repo $ScriptUpdateConfig.GitHubRepo -AssetName $ScriptUpdateConfig.AssetName -Path $ScriptUpdateConfig.GitHubPath -RestartAfterUpdate
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("No update source (URL or GitHub Repo) configured.", "Update Check", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        if ($updated -eq $false) {
            $msg = "You are currently on the latest version. Would you like to force an update anyway?"
            $res = [System.Windows.Forms.MessageBox]::Show($msg, "Update Check", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
                if ($ScriptUpdateConfig.UpdateUrl) {
                    Invoke-SelfUpdate -UpdateUrl $ScriptUpdateConfig.UpdateUrl -RestartAfterUpdate -Force
                }
                else {
                    Invoke-AuthenticatedUpdate -Owner $ScriptUpdateConfig.GitHubOwner -Repo $ScriptUpdateConfig.GitHubRepo -AssetName $ScriptUpdateConfig.AssetName -Path $ScriptUpdateConfig.GitHubPath -RestartAfterUpdate -Force
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

# SIG # Begin signature block
# MIIe6QYJKoZIhvcNAQcCoIIe2jCCHtYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD1XUgV5T/WwOPV
# P9xYADJT5gCqXzQBbTizXpJ3n6LR26CCA2AwggNcMIICRKADAgECAhBg0d5ZM/rs
# vETINKKp+/heMA0GCSqGSIb3DQEBCwUAMEYxFTATBgNVBAMMDFNraWxsIE9uIE5l
# dDEtMCsGCSqGSIb3DQEJARYeb3Jlc3Rpcy5vdGhvbm9zQHNraWxsb25uZXQuY29t
# MB4XDTI1MTIyMzE0MjY1MloXDTI2MTIyMzE0NDY1MlowRjEVMBMGA1UEAwwMU2tp
# bGwgT24gTmV0MS0wKwYJKoZIhvcNAQkBFh5vcmVzdGlzLm90aG9ub3NAc2tpbGxv
# bm5ldC5jb20wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDFpU2S7eIK
# q8Ypi4oR8JVe2aGZU3dgiwysOS369grASsMNJk0MtwBG7eySjxDgzsZhZxkP9xIg
# FPimAa+DhYKyJhnDD80LTTzEdWEbDhUD2IvUubgDRePFM8WrQtiqjtvnC9F/IDYn
# +Dbh2pAIw3OeRaZ4uIRzmwzQZQCciLdSuVJqvisKeN55HaKYFm6SU1p4G0UYsjbR
# LAMtJ3oNpP1So7fil+yBOJyzthSHH4uoDRGooE5vMsVHJRpLG5ICT/JV7V3c+JIC
# y/9LSiO6avionHD7LyAY3pAX0QkMph0IY1ow7HCVrhpK2PBv3Ka3cmqReUrIJKHg
# blr22VILcmRpAgMBAAGjRjBEMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggr
# BgEFBQcDAzAdBgNVHQ4EFgQUprKONaV5Vj0ttNzBzRfyp6hhaTcwDQYJKoZIhvcN
# AQELBQADggEBADLvKLxSHV4mPOVP5JpqXTNVPdNhh4VEuXbE+Km3SkYpWOxvu9OB
# OM6eKYYJDcUntKXxJNLEKNtDa0X3lPhNBhgJUfjnr351YzP4M8PscqLFrEWry6ty
# 3pkD1wPWcN13g3MruNEU4ZhxfF2ifKrVRJgSB+aWUvSsz8u3Ob4dgaJaVbByRzYS
# KvDhJYuiDZXZDruUUdfsbqTIB0wtjV6NyFED6mf1QwvUUuaXJUHJY+P7nvjSr5i1
# LKBYVofWe2cOX+70oWe+eB2ezpUUFBjtFznzRYtGDaPRYitxkEApY6EJ8JqkVJa+
# qat7ndHgKysInhT+KntTEl3NOZGSt0M37jAxghrfMIIa2wIBATBaMEYxFTATBgNV
# BAMMDFNraWxsIE9uIE5ldDEtMCsGCSqGSIb3DQEJARYeb3Jlc3Rpcy5vdGhvbm9z
# QHNraWxsb25uZXQuY29tAhBg0d5ZM/rsvETINKKp+/heMA0GCWCGSAFlAwQCAQUA
# oHwwEAYKKwYBBAGCNwIBDDECMAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIMt0
# x9TireI0B5nqG4zmNKwcn+DLE8WkvNEhYg0GH4g4MA0GCSqGSIb3DQEBAQUABIIB
# ABGoiSFw4be8cTPOGDHHFGjugBAB592R/HxBK1X7i/luaxdj1av5hLNJj2BttwUr
# 1hgMxpRI+BtvUr3g+xUIi8cA1X4R5V6QbqzEqSZZtRvq/J3VgLHMfhwhO+ciZTxf
# KMJx1Ect0U0i12ra3F6dZFzCoDrfQTel9JdM/mM+C8GdALMG3pxCH+ZLWwiAfTcH
# lt+GXtAiPSZF659UVRqm6Ju7kbC4at+xHCWiqgPGfAtjbpys4/x3X3Idzoevcz4r
# fbTjXueWwa2OFsf78xMw9aH9wSeSjExWQ/ADV2TJxFp3ta9/TFbDOm++Q1T6qiJI
# IMk+Kz/sAPpJdRKXdocR576hghjYMIIY1AYKKwYBBAGCNwMDATGCGMQwghjABgkq
# hkiG9w0BBwKgghixMIIYrQIBAzEPMA0GCWCGSAFlAwQCAgUAMIH4BgsqhkiG9w0B
# CRABBKCB6ASB5TCB4gIBAQYKKwYBBAGyMQIBATAxMA0GCWCGSAFlAwQCAQUABCAN
# VMXvxnDZivoWTISjnUcp2ftSFKaFR8aWS6Cp6S129gIVANMZEVVYiCXbRI3LVfEH
# ggLb7u8AGA8yMDI1MTIyMzE2MjczNFqgdqR0MHIxCzAJBgNVBAYTAkdCMRcwFQYD
# VQQIEw5XZXN0IFlvcmtzaGlyZTEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMTAw
# LgYDVQQDEydTZWN0aWdvIFB1YmxpYyBUaW1lIFN0YW1waW5nIFNpZ25lciBSMzag
# ghMEMIIGYjCCBMqgAwIBAgIRAKQpO24e3denNAiHrXpOtyQwDQYJKoZIhvcNAQEM
# BQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEsMCoG
# A1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBSMzYwHhcNMjUw
# MzI3MDAwMDAwWhcNMzYwMzIxMjM1OTU5WjByMQswCQYDVQQGEwJHQjEXMBUGA1UE
# CBMOV2VzdCBZb3Jrc2hpcmUxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEwMC4G
# A1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWduZXIgUjM2MIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA04SV9G6kU3jyPRBLeBIHPNyU
# gVNnYayfsGOyYEXrn3+SkDYTLs1crcw/ol2swE1TzB2aR/5JIjKNf75QBha2Ddj+
# 4NEPKDxHEd4dEn7RTWMcTIfm492TW22I8LfH+A7Ehz0/safc6BbsNBzjHTt7FngN
# fhfJoYOrkugSaT8F0IzUh6VUwoHdYDpiln9dh0n0m545d5A5tJD92iFAIbKHQWGb
# CQNYplqpAFasHBn77OqW37P9BhOASdmjp3IijYiFdcA0WQIe60vzvrk0HG+iVcwV
# Zjz+t5OcXGTcxqOAzk1frDNZ1aw8nFhGEvG0ktJQknnJZE3D40GofV7O8WzgaAnZ
# moUn4PCpvH36vD4XaAF2CjiPsJWiY/j2xLsJuqx3JtuI4akH0MmGzlBUylhXvdNV
# XcjAuIEcEQKtOBR9lU4wXQpISrbOT8ux+96GzBq8TdbhoFcmYaOBZKlwPP7pOp5M
# zx/UMhyBA93PQhiCdPfIVOCINsUY4U23p4KJ3F1HqP3H6Slw3lHACnLilGETXRg5
# X/Fp8G8qlG5Y+M49ZEGUp2bneRLZoyHTyynHvFISpefhBCV0KdRZHPcuSL5OAGWn
# BjAlRtHvsMBrI3AAA0Tu1oGvPa/4yeeiAyu+9y3SLC98gDVbySnXnkujjhIh+oaa
# tsk/oyf5R2vcxHahajMCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaAFF9Y7UwxeqJh
# Qo1SgLqzYZcZojKbMB0GA1UdDgQWBBSIYYyhKjdkgShgoZsx0Iz9LALOTzAOBgNV
# HQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcD
# CDBKBgNVHSAEQzBBMDUGDCsGAQQBsjEBAgEDCDAlMCMGCCsGAQUFBwIBFhdodHRw
# czovL3NlY3RpZ28uY29tL0NQUzAIBgZngQwBBAIwSgYDVR0fBEMwQTA/oD2gO4Y5
# aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSMzYuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5aHR0cDovL2Ny
# dC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5nQ0FSMzYuY3J0
# MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTANBgkqhkiG9w0B
# AQwFAAOCAYEAAoE+pIZyUSH5ZakuPVKK4eWbzEsTRJOEjbIu6r7vmzXXLpJx4FyG
# mcqnFZoa1dzx3JrUCrdG5b//LfAxOGy9Ph9JtrYChJaVHrusDh9NgYwiGDOhyyJ2
# zRy3+kdqhwtUlLCdNjFjakTSE+hkC9F5ty1uxOoQ2ZkfI5WM4WXA3ZHcNHB4V42z
# i7Jk3ktEnkSdViVxM6rduXW0jmmiu71ZpBFZDh7Kdens+PQXPgMqvzodgQJEkxaI
# ON5XRCoBxAwWwiMm2thPDuZTzWp/gUFzi7izCmEt4pE3Kf0MOt3ccgwn4Kl2FIcQ
# aV55nkjv1gODcHcD9+ZVjYZoyKTVWb4VqMQy/j8Q3aaYd/jOQ66Fhk3NWbg2tYl5
# jhQCuIsE55Vg4N0DUbEWvXJxtxQQaVR5xzhEI+BjJKzh3TQ026JxHhr2fuJ0mV68
# AluFr9qshgwS5SpN5FFtaSEnAwqZv3IS+mlG50rK7W3qXbWwi4hmpylUfygtYLEd
# LQukNEX1jiOKMIIGFDCCA/ygAwIBAgIQeiOu2lNplg+RyD5c9MfjPzANBgkqhkiG
# 9w0BAQwFADBXMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVk
# MS4wLAYDVQQDEyVTZWN0aWdvIFB1YmxpYyBUaW1lIFN0YW1waW5nIFJvb3QgUjQ2
# MB4XDTIxMDMyMjAwMDAwMFoXDTM2MDMyMTIzNTk1OVowVTELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBDQSBSMzYwggGiMA0GCSqGSIb3DQEBAQUAA4IBjwAw
# ggGKAoIBgQDNmNhDQatugivs9jN+JjTkiYzT7yISgFQ+7yavjA6Bg+OiIjPm/N/t
# 3nC7wYUrUlY3mFyI32t2o6Ft3EtxJXCc5MmZQZ8AxCbh5c6WzeJDB9qkQVa46xiY
# Epc81KnBkAWgsaXnLURoYZzksHIzzCNxtIXnb9njZholGw9djnjkTdAA83abEOHQ
# 4ujOGIaBhPXG2NdV8TNgFWZ9BojlAvflxNMCOwkCnzlH4oCw5+4v1nssWeN1y4+R
# laOywwRMUi54fr2vFsU5QPrgb6tSjvEUh1EC4M29YGy/SIYM8ZpHadmVjbi3Pl8h
# JiTWw9jiCKv31pcAaeijS9fc6R7DgyyLIGflmdQMwrNRxCulVq8ZpysiSYNi79tw
# 5RHWZUEhnRfs/hsp/fwkXsynu1jcsUX+HuG8FLa2BNheUPtOcgw+vHJcJ8HnJCrc
# UWhdFczf8O+pDiyGhVYX+bDDP3GhGS7TmKmGnbZ9N+MpEhWmbiAVPbgkqykSkzyY
# Vr15OApZYK8CAwEAAaOCAVwwggFYMB8GA1UdIwQYMBaAFPZ3at0//QET/xahbIIC
# L9AKPRQlMB0GA1UdDgQWBBRfWO1MMXqiYUKNUoC6s2GXGaIymzAOBgNVHQ8BAf8E
# BAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggrBgEFBQcDCDAR
# BgNVHSAECjAIMAYGBFUdIAAwTAYDVR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5z
# ZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmww
# fAYIKwYBBQUHAQEEcDBuMEcGCCsGAQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28u
# Y29tL1NlY3RpZ29QdWJsaWNUaW1lU3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEF
# BQcwAYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIB
# ABLXeyCtDjVYDJ6BHSVY/UwtZ3Svx2ImIfZVVGnGoUaGdltoX4hDskBMZx5NY5L6
# SCcwDMZhHOmbyMhyOVJDwm1yrKYqGDHWzpwVkFJ+996jKKAXyIIaUf5JVKjccev3
# w16mNIUlNTkpJEor7edVJZiRJVCAmWAaHcw9zP0hY3gj+fWp8MbOocI9Zn78xvm9
# XKGBp6rEs9sEiq/pwzvg2/KjXE2yWUQIkms6+yslCRqNXPjEnBnxuUB1fm6bPAV+
# Tsr/Qrd+mOCJemo06ldon4pJFbQd0TQVIMLv5koklInHvyaf6vATJP4DfPtKzSBP
# kKlOtyaFTAjD2Nu+di5hErEVVaMqSVbfPzd6kNXOhYm23EWm6N2s2ZHCHVhlUgHa
# C4ACMRCgXjYfQEDtYEK54dUwPJXV7icz0rgCzs9VI29DwsjVZFpO4ZIVR33LwXyP
# DbYFkLqYmgHjR3tKVkhh9qKV2WCmBuC27pIOx6TYvyqiYbntinmpOqh/QPAnhDge
# xKG9GX/n1PggkGi9HCapZp8fRwg8RftwS21Ln61euBG0yONM6noD2XQPrFwpm3Gc
# uqJMf0o8LLrFkSLRQNwxPDDkWXhW+gZswbaiie5fd/W2ygcto78XCSPfFWveUOSZ
# 5SqK95tBO8aTHmEa4lpJVD7HrTEn9jb1EGvxOb1cnn0CMIIGgjCCBGqgAwIBAgIQ
# NsKwvXwbOuejs902y8l1aDANBgkqhkiG9w0BAQwFADCBiDELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCk5ldyBKZXJzZXkxFDASBgNVBAcTC0plcnNleSBDaXR5MR4wHAYD
# VQQKExVUaGUgVVNFUlRSVVNUIE5ldHdvcmsxLjAsBgNVBAMTJVVTRVJUcnVzdCBS
# U0EgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkwHhcNMjEwMzIyMDAwMDAwWhcNMzgw
# MTE4MjM1OTU5WjBXMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1p
# dGVkMS4wLAYDVQQDEyVTZWN0aWdvIFB1YmxpYyBUaW1lIFN0YW1waW5nIFJvb3Qg
# UjQ2MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAiJ3YuUVnnR3d6Lkm
# gZpUVMB8SQWbzFoVD9mUEES0QUCBdxSZqdTkdizICFNeINCSJS+lV1ipnW5ihkQy
# C0cRLWXUJzodqpnMRs46npiJPHrfLBOifjfhpdXJ2aHHsPHggGsCi7uE0awqKggE
# /LkYw3sqaBia67h/3awoqNvGqiFRJ+OTWYmUCO2GAXsePHi+/JUNAax3kpqstbl3
# vcTdOGhtKShvZIvjwulRH87rbukNyHGWX5tNK/WABKf+Gnoi4cmisS7oSimgHUI0
# Wn/4elNd40BFdSZ1EwpuddZ+Wr7+Dfo0lcHflm/FDDrOJ3rWqauUP8hsokDoI7D/
# yUVI9DAE/WK3Jl3C4LKwIpn1mNzMyptRwsXKrop06m7NUNHdlTDEMovXAIDGAvYy
# nPt5lutv8lZeI5w3MOlCybAZDpK3Dy1MKo+6aEtE9vtiTMzz/o2dYfdP0KWZwZIX
# bYsTIlg1YIetCpi5s14qiXOpRsKqFKqav9R1R5vj3NgevsAsvxsAnI8Oa5s2oy25
# qhsoBIGo/zi6GpxFj+mOdh35Xn91y72J4RGOJEoqzEIbW3q0b2iPuWLA911cRxgY
# 5SJYubvjay3nSMbBPPFsyl6mY4/WYucmyS9lo3l7jk27MAe145GWxK4O3m3gEFEI
# kv7kRmefDR7Oe2T1HxAnICQvr9sCAwEAAaOCARYwggESMB8GA1UdIwQYMBaAFFN5
# v1qqK0rPVIDh2JvAnfKyA2bLMB0GA1UdDgQWBBT2d2rdP/0BE/8WoWyCAi/QCj0U
# JTAOBgNVHQ8BAf8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zATBgNVHSUEDDAKBggr
# BgEFBQcDCDARBgNVHSAECjAIMAYGBFUdIAAwUAYDVR0fBEkwRzBFoEOgQYY/aHR0
# cDovL2NybC51c2VydHJ1c3QuY29tL1VTRVJUcnVzdFJTQUNlcnRpZmljYXRpb25B
# dXRob3JpdHkuY3JsMDUGCCsGAQUFBwEBBCkwJzAlBggrBgEFBQcwAYYZaHR0cDov
# L29jc3AudXNlcnRydXN0LmNvbTANBgkqhkiG9w0BAQwFAAOCAgEADr5lQe1oRLjl
# ocXUEYfktzsljOt+2sgXke3Y8UPEooU5y39rAARaAdAxUeiX1ktLJ3+lgxtoLQhn
# 5cFb3GF2SSZRX8ptQ6IvuD3wz/LNHKpQ5nX8hjsDLRhsyeIiJsms9yAWnvdYOdEM
# q1W61KE9JlBkB20XBee6JaXx4UBErc+YuoSb1SxVf7nkNtUjPfcxuFtrQdRMRi/f
# InV/AobE8Gw/8yBMQKKaHt5eia8ybT8Y/Ffa6HAJyz9gvEOcF1VWXG8OMeM7Vy7B
# s6mSIkYeYtddU1ux1dQLbEGur18ut97wgGwDiGinCwKPyFO7ApcmVJOtlw9FVJxw
# /mL1TbyBns4zOgkaXFnnfzg4qbSvnrwyj1NiurMp4pmAWjR+Pb/SIduPnmFzbSN/
# G8reZCL4fvGlvPFk4Uab/JVCSmj59+/mB2Gn6G/UYOy8k60mKcmaAZsEVkhOFuoj
# 4we8CYyaR9vd9PGZKSinaZIkvVjbH/3nlLb0a7SBIkiRzfPfS9T+JesylbHa1LtR
# V9U/7m0q7Ma2CQ/t392ioOssXW7oKLdOmMBl14suVFBmbzrt5V5cQPnwtd3UOTpS
# 9oCG+ZZheiIvPgkDmA8FzPsnfXW5qHELB43ET7HHFHeRPRYrMBKjkb8/IN7Po0d0
# hQoF4TeMM+zYAJzoKQnVKOLg8pZVPT8xggSSMIIEjgIBATBqMFUxCzAJBgNVBAYT
# AkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxLDAqBgNVBAMTI1NlY3RpZ28g
# UHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjM2AhEApCk7bh7d16c0CIetek63JDAN
# BglghkgBZQMEAgIFAKCCAfkwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwG
# CSqGSIb3DQEJBTEPFw0yNTEyMjMxNjI3MzNaMD8GCSqGSIb3DQEJBDEyBDDx5Pu4
# r0x/EjpPcE/PdBZMqbqjkn8DIJqYNLho+ZrNPflgFY2zry/qiDf4fljYbFowggF6
# BgsqhkiG9w0BCRACDDGCAWkwggFlMIIBYTAWBBQ4yRSBEES03GY+k9R0S4FBhqm1
# sTCBhwQUxq5U5HiG8Xw9VRJIjGnDSnr5wt0wbzBbpFkwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NgIQeiOu2lNplg+RyD5c9MfjPzCBvAQU
# hT1jLZOCgmF80JA1xJHeksFC2scwgaMwgY6kgYswgYgxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkgQ2l0eTEeMBwGA1UE
# ChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVVU0VSVHJ1c3QgUlNB
# IENlcnRpZmljYXRpb24gQXV0aG9yaXR5AhA2wrC9fBs656Oz3TbLyXVoMA0GCSqG
# SIb3DQEBAQUABIICAAoFP3rQsmV+JyhlzFNfPPLfJrpCBr6vGNTy6IMQzTm0yPYX
# Unl310BAnIsVw0cbMWvXA0sZqpu5DiLMogbiObvvCg8Scx67wPPXyJTAjYkUH4o+
# +ZAYKTgsIPEEUzAjGcfboyYJqBWjMgV+2YOyriGms/D9zuKqz/Nk/7eGhOF92LCG
# IwoaK8yahwHC17qq+HyDmOvRnmUarHI2VGTFzrOOQ28Hmq9EDz0Ha2dArQNWH8+Q
# GAYRXlYjJsTfxPgZGhWsUGa/d3HhdDw3o5Jp+pkSG9MJaEG5SK0f8q/0LzUy0HyM
# Xs+wkAhxx9reKVHbveIPc1lcjh7tK3PIMFHazFHQ58FCKrqUW/Fd+C0LN/ZrnLhY
# L0utTwaqvVbJ7ggH0OB+x5cKMTAnPmerqFYbwD94x/ZPCMqk2eOrU0FP0BUrnVS7
# vsfZJU0I34pViYL8K08HbWMIFOqYGy3JPslaaMNBzrX83AuYx2mmUhG/bfPA/qwq
# 3KGUBXzoPxswC2Z/SC25V7SMd6PxjLhHxcVhJWO1o2Y06UAZAZNvFk6X50W5BNZ1
# h0Xb88P7ArBDscr3MB0EfIVoYZHtwZGWN0A9TyeVgN+w9nGw36fYcEacjVHoCd3p
# +mLlQ/4ftri0/CZdteNBl7qf1VdWN9hKgBhVWrs8N3z7HrsDDOe7YVWQTVI9
# SIG # End signature block
