# Define the network path
$networkPath = "\\filesrv3.kpax.local\es-kam"

# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

function Get-ModernCredential {
    # Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Authentication Required"
    $form.Size = New-Object System.Drawing.Size(400, 320)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::White
    $form.TopMost = $true

    # Header
    $header = New-StyledLabel "Network Authentication" 20 $true
    $form.Controls.Add($header)

    $subHeader = New-StyledLabel "Please enter your credentials for the network drive." 50
    $subHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($subHeader)

    # Label: Username
    $labelUser = New-StyledLabel "Username" 90
    $form.Controls.Add($labelUser)

    # TextBox: Username
    $textBoxUser = New-Object System.Windows.Forms.TextBox
    $textBoxUser.Location = New-Object System.Drawing.Point(25, 115)
    $textBoxUser.Size = New-Object System.Drawing.Size(330, 26)
    $textBoxUser.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    # Default to current user if desired, or leave empty
    $textBoxUser.Text = $env:USERNAME
    $form.Controls.Add($textBoxUser)

    # Label: Password
    $labelPassword = New-StyledLabel "Password" 155
    $form.Controls.Add($labelPassword)

    # TextBox: Password
    $textBoxPassword = New-Object System.Windows.Forms.TextBox
    $textBoxPassword.Location = New-Object System.Drawing.Point(25, 180)
    $textBoxPassword.Size = New-Object System.Drawing.Size(330, 26)
    $textBoxPassword.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $textBoxPassword.PasswordChar = "●"
    $form.Controls.Add($textBoxPassword)

    # Separator line
    $line = New-Object System.Windows.Forms.Label
    $line.Location = New-Object System.Drawing.Point(0, 230)
    $line.Size = New-Object System.Drawing.Size(400, 1)
    $line.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $form.Controls.Add($line)

    # Button Panel Background
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Location = New-Object System.Drawing.Point(0, 231)
    $buttonPanel.Size = New-Object System.Drawing.Size(400, 50)
    $buttonPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $form.Controls.Add($buttonPanel)

    # Button: OK (Connect)
    $okButton = New-StyledButton "Connect" 170 8 $true
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $buttonPanel.Controls.Add($okButton)

    # Button: Cancel
    $cancelButton = New-StyledButton "Cancel" 280 8 $false
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $buttonPanel.Controls.Add($cancelButton)

    # Show the dialog
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $user = $textBoxUser.Text
        if ($user -notlike "*@*") {
            $user = "$user@kpax.local"
        }
        $pass = $textBoxPassword.Text
        $securePass = ConvertTo-SecureString $pass -AsPlainText -Force
        return New-Object System.Management.Automation.PSCredential ($user, $securePass)
    }
    
    return $null
}

function Show-SuccessPopup ($driveLetter) {
    $successForm = New-Object System.Windows.Forms.Form
    $successForm.Text = "Success"
    $successForm.Size = New-Object System.Drawing.Size(400, 180)
    $successForm.StartPosition = "CenterScreen"
    $successForm.FormBorderStyle = "FixedDialog"
    $successForm.MaximizeBox = $false
    $successForm.MinimizeBox = $false
    $successForm.TopMost = $true
    $successForm.BackColor = [System.Drawing.Color]::White

    $lblSuccess = New-StyledLabel "Drive mapped successfully!" 20 $true
    $lblSuccess.ForeColor = [System.Drawing.Color]::SeaGreen
    $successForm.Controls.Add($lblSuccess)
    
    $lblInfo = New-StyledLabel "Mapped to drive letter: $driveLetter" 60
    $successForm.Controls.Add($lblInfo)
    
    $btnOpen = New-StyledButton "Open Drive" 70 100 $true
    $btnOpen.Width = 120
    $btnOpen.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $successForm.Controls.Add($btnOpen)
    
    $btnClose = New-StyledButton "Close" 210 100 $false
    $btnClose.Width = 120
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::No
    $successForm.Controls.Add($btnClose)

    $result = $successForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        explorer.exe "$driveLetter"
    }
}

function Show-ErrorPopup ($message) {
    $errForm = New-Object System.Windows.Forms.Form
    $errForm.Text = "Error"
    $errForm.Size = New-Object System.Drawing.Size(400, 180)
    $errForm.StartPosition = "CenterScreen"
    $errForm.FormBorderStyle = "FixedDialog"
    $errForm.MaximizeBox = $false
    $errForm.MinimizeBox = $false
    $errForm.TopMost = $true
    $errForm.BackColor = [System.Drawing.Color]::White

    $lblErr = New-StyledLabel "Connection Failed" 20 $true
    $lblErr.ForeColor = [System.Drawing.Color]::Firebrick
    $errForm.Controls.Add($lblErr)
    
    $lblErrMsg = New-StyledLabel $message 60
    $lblErrMsg.Size = New-Object System.Drawing.Size(350, 40)
    $errForm.Controls.Add($lblErrMsg)
    
    $btnClose = New-StyledButton "Close" 145 90 $true
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $errForm.Controls.Add($btnClose)
    $errForm.AcceptButton = $btnClose
    
    $errForm.ShowDialog() | Out-Null
}

# Get all used drive letters
$usedDrives = (Get-PSDrive -PSProvider FileSystem).Name

# Define the alphabet range for drive letters
$alphabet = 65..90 | ForEach-Object { [char]$_ }

# Find the first available drive letter
$availableDrive = $alphabet | Where-Object { $_ -notin $usedDrives } | Select-Object -First 1

# Check if the drive is already mapped
$cleanNetworkPath = $networkPath -replace '\\+', '\'
$existingDrive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType = 4" | Where-Object { 
    $pName = $_.ProviderName -replace '\\+', '\'
    $pName -ieq $cleanNetworkPath
}

if ($existingDrive) {
    Write-Host "The network path $networkPath is already mapped to drive $($existingDrive.DeviceID)" -ForegroundColor White
    
    # Show notice popup
    $noticeForm = New-Object System.Windows.Forms.Form
    $noticeForm.Text = "Already Connected"
    $noticeForm.Size = New-Object System.Drawing.Size(400, 180)
    $noticeForm.StartPosition = "CenterScreen"
    $noticeForm.FormBorderStyle = "FixedDialog"
    $noticeForm.MaximizeBox = $false
    $noticeForm.MinimizeBox = $false
    $noticeForm.TopMost = $true
    $noticeForm.BackColor = [System.Drawing.Color]::White

    $lblNotice = New-StyledLabel "Drive Already Mapped" 20 $true
    $lblNotice.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $noticeForm.Controls.Add($lblNotice)
    
    $lblMsgText = "The folder is already mapped to drive letter: $($existingDrive.DeviceID)"
    $lblMsg = New-StyledLabel $lblMsgText 60
    $lblMsg.Size = New-Object System.Drawing.Size(350, 40)
    $noticeForm.Controls.Add($lblMsg)
    
    $btnOpen = New-StyledButton "Open Drive" 70 100 $true
    $btnOpen.Width = 120
    $btnOpen.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $noticeForm.Controls.Add($btnOpen)
    
    $btnExit = New-StyledButton "Close" 210 100 $false
    $btnExit.Width = 120
    $btnExit.DialogResult = [System.Windows.Forms.DialogResult]::No
    $noticeForm.Controls.Add($btnExit)
    
    $result = $noticeForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        explorer.exe "$($existingDrive.DeviceID)"
    }
    exit
}

if ($availableDrive) {
    $driveMapped = $false
    
    # 1. Try to map with current credentials (Silent attempt)
    try {
        Write-Host "Attempting to map using current credentials..." -ForegroundColor Cyan
        New-PSDrive -Name $availableDrive -PSProvider FileSystem -Root $networkPath -Persist -ErrorAction Stop | Out-Null
        Write-Host "Successfully mapped $networkPath to drive $availableDrive (Cached/Current Credentials)" -ForegroundColor Green
        Show-SuccessPopup $availableDrive
        $driveMapped = $true
    }
    catch {
        Write-Host "Silent mapping failed. Prompting for credentials..." -ForegroundColor Yellow
    }

    # 2. If silent map failed, prompt user
    if (-not $driveMapped) {
        $credential = Get-ModernCredential

        if ($credential) {
            # Map the network drive with provided credentials
            try {
                New-PSDrive -Name $availableDrive -PSProvider FileSystem -Root $networkPath -Credential $credential -Persist -ErrorAction Stop | Out-Null
                Write-Host "Mapped $networkPath to drive $availableDrive" -ForegroundColor Green
                Show-SuccessPopup $availableDrive
            }
            catch {
                Write-Error "Failed to map drive: $_"
                Show-ErrorPopup "Check your credentials and try again."
            }
        }
        else {
            Write-Warning "Setup cancelled by user."
        }
    }
}
else {
    Write-Host "No available drive letters."
    Show-ErrorPopup "No available drive letters found."
}
