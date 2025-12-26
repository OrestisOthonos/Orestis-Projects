<#
.SYNOPSIS
Deletes orphaned ("Account Unknown") user profiles.

.DESCRIPTION
Lists local user profiles whose SIDs cannot be resolved to accounts and
allows selection and deletion. Uses Out-GridView when available with a
WinForms fallback. Supports `-WhatIf` and `-Confirm` via `ShouldProcess`.

.NOTES
Run PowerShell as Administrator to allow deletion of profiles.
#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [switch]$UseUI,
    [switch]$NoUI
)

Set-StrictMode -Version Latest

function Write-Log {
    param([Parameter(Mandatory=$true)]$Message)
    try {
        $logPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'DeleteAccountProfiles_error.log' } else { Join-Path $env:TEMP 'DeleteAccountProfiles_error.log' }
        $time = (Get-Date).ToString('s')
        "$time`t$Message" | Out-File -FilePath $logPath -Encoding utf8 -Append
    } catch {}
}

function Test-SidResolution {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$SID)
    try {
        $sidObj = [System.Security.Principal.SecurityIdentifier]::new($SID)
        $null = $sidObj.Translate([System.Security.Principal.NTAccount])
        return $true
    } catch {
        return $false
    }
}

# Load WinForms types when UI is requested
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

function Show-ModernUI {
    param(
        [Parameter(Mandatory=$false)][array]$Items
    )
    if (-not $Items) { $Items = @() }
    [void][System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'User Profile Cleanup'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(900,560)
    $form.Font = New-Object System.Drawing.Font('Segoe UI',9)
    # Make the window non-resizable and disable maximize
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    # Modern styling
    $accent = [System.Drawing.Color]::FromArgb(0,102,204)
    $bg = [System.Drawing.Color]::FromArgb(245,247,250)
    $form.BackColor = $bg
    try { $form.Icon = [System.Drawing.SystemIcons]::Application } catch {}

    $tl = New-Object System.Windows.Forms.TableLayoutPanel
    $tl.Dock = 'Fill'
    $tl.ColumnCount = 1
    $tl.RowCount = 3
    $tl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,40)))
    $tl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
    $tl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,50)))
    $form.Controls.Add($tl)

    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = 'Fill'
    $topPanel.BackColor = $bg
    $tl.Controls.Add($topPanel,0,0)

    # Title label
    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'User Profile Cleanup'
    $title.Font = New-Object System.Drawing.Font('Segoe UI',11,[System.Drawing.FontStyle]::Bold)
    $title.AutoSize = $true
    $title.Location = New-Object System.Drawing.Point(10,6)
    $topPanel.Controls.Add($title)

    $search = New-Object System.Windows.Forms.TextBox
    # PlaceholderText may not be available in all runtimes (ps2exe/exe); avoid using it
    $search.Text = ''
    $search.Width = 420
    $search.Location = New-Object System.Drawing.Point(10,8)
    $topPanel.Controls.Add($search)

    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = 'Select All'
    $btnSelectAll.Location = New-Object System.Drawing.Point(450,6)
    $btnSelectAll.Width = 90
    $btnSelectAll.FlatStyle = 'Flat'
    $btnSelectAll.BackColor = $bg
    $btnSelectAll.ForeColor = [System.Drawing.Color]::Black
    $topPanel.Controls.Add($btnSelectAll)

    $btnInvert = New-Object System.Windows.Forms.Button
    $btnInvert.Text = 'Invert'
    $btnInvert.Location = New-Object System.Drawing.Point(550,6)
    $btnInvert.Width = 70
    $btnInvert.FlatStyle = 'Flat'
    $btnInvert.BackColor = $bg
    $btnInvert.ForeColor = [System.Drawing.Color]::Black
    $topPanel.Controls.Add($btnInvert)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = 'Refresh'
    $btnRefresh.Location = New-Object System.Drawing.Point(630,6)
    $btnRefresh.Width = 80
    $btnRefresh.FlatStyle = 'Flat'
    $btnRefresh.BackColor = $bg
    $btnRefresh.ForeColor = [System.Drawing.Color]::Black
    $topPanel.Controls.Add($btnRefresh)

    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Dock = 'Fill'
    $dgv.AllowUserToAddRows = $false
    $dgv.AutoGenerateColumns = $false
    $dgv.RowHeadersVisible = $false
    $dgv.SelectionMode = 'FullRowSelect'
    # Modern grid styling
    $dgv.EnableHeadersVisualStyles = $false
    $dgv.BackgroundColor = [System.Drawing.Color]::FromArgb(238,241,245)
    $dgv.GridColor = [System.Drawing.Color]::FromArgb(220,224,230)
    $dgv.DefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI',9)
    $dgv.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,255,255)
    $dgv.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245,250,255)
    $dgv.RowTemplate.Height = 26
    $dgv.ColumnHeadersDefaultCellStyle.BackColor = $accent
    $dgv.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $dgv.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
    $dgv.ColumnHeadersHeightSizeMode = 'AutoSize'

    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheck.HeaderText = 'Delete'
    $colCheck.DataPropertyName = 'Delete'
    $colCheck.Width = 60
    $colCheck.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
    $dgv.Columns.Add($colCheck) | Out-Null

    $colPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPath.HeaderText = 'Profile Path'
    $colPath.DataPropertyName = 'Path'
    $colPath.Width = 520
    $colPath.ReadOnly = $true
    $colPath.MinimumWidth = 200
    $colPath.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $colPath.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,255,255)
    $dgv.Columns.Add($colPath) | Out-Null

    $colExists = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colExists.HeaderText = 'Exists'
    $colExists.DataPropertyName = 'Exists'
    $colExists.Width = 60
    $colExists.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
    $colExists.ReadOnly = $true
    $colExists.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
    $dgv.Columns.Add($colExists) | Out-Null

    $colSid = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSid.HeaderText = 'SID'
    $colSid.DataPropertyName = 'SID'
    $colSid.Width = 240
    $colSid.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
    $colSid.ReadOnly = $true
    $colSid.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,255,255)
    $dgv.Columns.Add($colSid) | Out-Null

    $colOrphan = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colOrphan.HeaderText = 'Orphan'
    $colOrphan.DataPropertyName = 'Orphan'
    $colOrphan.Width = 80
    $colOrphan.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
    $colOrphan.ReadOnly = $true
    $colOrphan.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250,250,250)
    $dgv.Columns.Add($colOrphan) | Out-Null

    $tl.Controls.Add($dgv,0,1)

    # Bottom area: status on left, centered buttons on right
    $bottom = New-Object System.Windows.Forms.TableLayoutPanel
    $bottom.Dock = 'Fill'
    $bottom.ColumnCount = 3
    $bottom.RowCount = 1
    $bottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
    $bottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $bottom.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
    $tl.Controls.Add($bottom,0,2)

    $status = New-Object System.Windows.Forms.Label
    $status.AutoSize = $true
    $status.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
    $bottom.Controls.Add($status,0,0)

    $btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnPanel.FlowDirection = 'LeftToRight'
    $btnPanel.WrapContents = $false
    $btnPanel.AutoSize = $true
    $btnPanel.AutoSizeMode = 'GrowOnly'
    $btnPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $bottom.Controls.Add($btnPanel,1,0)

    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = 'Delete Selected'
    $btnDelete.Width = 120
    $btnDelete.Height = 30
    $btnDelete.FlatStyle = 'Flat'
    $btnDelete.BackColor = [System.Drawing.Color]::FromArgb(220,53,69)
    $btnDelete.ForeColor = [System.Drawing.Color]::White
    $btnPanel.Controls.Add($btnDelete)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Width = 90
    $btnClose.Height = 30
    $btnClose.FlatStyle = 'Flat'
    $btnPanel.Controls.Add($btnClose)

    # Tooltips
    $tt = New-Object System.Windows.Forms.ToolTip
    $tt.SetToolTip($btnRefresh,'Refresh profile list')
    $tt.SetToolTip($btnSelectAll,'Select all visible rows')
    $tt.SetToolTip($btnInvert,'Invert selection')
    $tt.SetToolTip($btnDelete,'Delete selected profiles')

    # Prepare items with Delete and Exists properties
    $all = $Items | ForEach-Object {
        [PSCustomObject]@{
            Delete = ($_.Orphan -eq 'Yes')
            Path = $_.Path
            SID = $_.SID
            Exists = if ($_.Path -and (Test-Path -Path $_.Path -PathType Container)) { 'Yes' } else { 'No' }
            Orphan = $_.Orphan
            ProfileObj = $_.ProfileObj
        }
    }

    $bindingList = New-Object System.ComponentModel.BindingList[object]
    foreach ($it in $all) { $bindingList.Add($it) | Out-Null }
    $dgv.DataSource = $bindingList

    # Auto-size columns to content and adjust form width to fit columns (respect screen bounds)
    try {
        $dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
        $dgv.AutoResizeColumns()
        $colsWidth = 0
        foreach ($c in $dgv.Columns) { $colsWidth += $c.Width }
        # add padding for row header and margins
        $desiredClientWidth = $colsWidth + 40
        $screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
        $maxClientWidth = $screenWidth - 80
        if ($desiredClientWidth -gt $maxClientWidth) { $desiredClientWidth = $maxClientWidth }
        if ($desiredClientWidth -lt 600) { $desiredClientWidth = 600 }
        # set form width to accommodate the desired client width plus non-client border
        $form.Width = $desiredClientWidth + ($form.Width - $form.ClientSize.Width)
    } catch {
        Write-Log "Auto-resize failed: $($_.Exception.Message)"
    }

    # Open profile folder in Explorer when user double-clicks a row
    $dgv.add_CellDoubleClick({
        param($sender,$e)
        if ($e.RowIndex -lt 0) { return }
        try {
            $row = $dgv.Rows[$e.RowIndex]
            $item = $row.DataBoundItem
            $path = $item.Path
            if ($path -and (Test-Path -Path $path -PathType Container)) {
                Start-Process explorer -ArgumentList $path
            } else {
                [System.Windows.Forms.MessageBox]::Show(("Folder not found or inaccessible: {0}" -f $path),'Info',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            }
        } catch {
        }
    })

    # Suppress default DataGridView error dialog and log errors instead
    $dgv.add_DataError({
        param($sender,$e)
        try { Write-Log ("DataGridView error: $($e.Exception.ToString())") } catch {}
        $e.ThrowException = $false
        $e.Cancel = $true
    })

    # Helper: refresh filtered view
    $searchFilter = { param($text)
        if ([string]::IsNullOrWhiteSpace($text)) {
            $items = $all
        } else {
            $t = $text.ToLower()
            $items = $all | Where-Object { ($_.Path -and $_.Path.ToLower().Contains($t)) -or ($_.SID -and $_.SID.ToLower().Contains($t)) }
        }
        $bindingList.Clear()
        foreach ($it in $items) { $bindingList.Add($it) | Out-Null }
    }

    $search.Add_TextChanged({ & $searchFilter $search.Text })

    $btnSelectAll.Add_Click({ foreach ($i in $bindingList) { $i.Delete = $true }; $dgv.Refresh() })
    $btnInvert.Add_Click({ foreach ($i in $bindingList) { $i.Delete = -not $i.Delete }; $dgv.Refresh() })
    $btnRefresh.Add_Click({
        # Rebuild mapped items to ensure 'Delete' and 'Exists' properties exist
        $all = $Items | ForEach-Object {
            [PSCustomObject]@{
                Delete = ($_.Orphan -eq 'Yes')
                Path = $_.Path
                SID = $_.SID
                Exists = if ($_.Path -and (Test-Path -Path $_.Path -PathType Container)) { 'Yes' } else { 'No' }
                Orphan = $_.Orphan
                ProfileObj = $_.ProfileObj
            }
        }
        & $searchFilter $search.Text
        $status.Text = "Refreshed: $($bindingList.Count) rows"
    })

    $btnClose.Add_Click({ $form.Close() })

    $btnDelete.Add_Click({
        $toRemove = @()
        foreach ($row in $bindingList) { if ($row.Delete) { $toRemove += $row } }
        if ($toRemove.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No profiles selected for deletion.','Info',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $msg = "Are you sure you want to permanently delete the selected $($toRemove.Count) profile(s)? This cannot be undone."
        $res = [System.Windows.Forms.MessageBox]::Show($msg,'Confirm Deletion',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $i = 0
        foreach ($it in $toRemove) {
            $i++
            $status.Text = "Deleting $i of $($toRemove.Count): $($it.Path)"
            $form.Refresh()
            if ($PSCmdlet.ShouldProcess($it.Path,'Remove user profile')) {
                try {
                    Remove-CimInstance -InputObject $it.ProfileObj -ErrorAction Stop
                    $bindingList.Remove($it) | Out-Null
                } catch {
                    $err = $_.Exception.Message
                    [System.Windows.Forms.MessageBox]::Show(("Failed to delete profile {0}: {1}" -f $it.Path, $err),'Error',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                }
            }
        }
        [System.Windows.Forms.MessageBox]::Show('Selected profiles deleted.','Done',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        $status.Text = 'Done.'
    })

    [void]$form.ShowDialog()
}

try {

    # Ensure elevation for deletions (best-effort check)
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Warning 'Not running as Administrator; deletion may fail.'
    }

# Gather profiles (exclude Special and active/loaded profiles)
$Profiles = Get-CimInstance -ClassName Win32_UserProfile | Where-Object { -not $_.Special -and -not $_.Loaded }
$ProfileList = $Profiles | ForEach-Object {
    [PSCustomObject]@{
        SID = $_.SID
        Path = $_.LocalPath
        Orphan = -not (Test-SidResolution -SID $_.SID)
        ProfileObj = $_
    }
}

if (-not $ProfileList) {
    Write-Host 'No local user profiles found to examine.' -ForegroundColor Yellow
    # continue so the UI can still open (shows an empty list)
    $ProfileList = @()
}

# Convert for grid/view use
$gridItems = $ProfileList | Select-Object SID, Path, @{Name='Exists';Expression={if ($_.Path -and (Test-Path -Path $_.Path -PathType Container)) {'Yes'} else {'No'}}}, @{Name='Orphan';Expression={if ($_.Orphan) {'Yes'} else {'No'}}}, ProfileObj

# By default prefer the modern WinForms UI. Use -NoUI to opt-out.
if (-not $NoUI) {
    try {
        # expose items to the new thread via a global variable so the STA thread can access them
        $global:__ShowModernUI_Items = $gridItems
        if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
            $t = [System.Threading.Thread]::new([System.Threading.ThreadStart]{ Show-ModernUI -Items $global:__ShowModernUI_Items })
            $t.SetApartmentState([System.Threading.ApartmentState]::STA)
            $t.Start()
            $t.Join()
            Remove-Variable -Name '__ShowModernUI_Items' -Scope Global -ErrorAction SilentlyContinue
        } else {
            Show-ModernUI -Items $gridItems
        }
        return
    } catch {
        $err = $_.Exception.Message
        Write-Log "Show-ModernUI failed: $err"
        Write-Warning "Modern UI failed to start; falling back to console/Out-GridView. See log for details."
    }
}

# If not explicitly requesting the UI, prefer Out-GridView when available
if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
    $selected = $gridItems | Where-Object Orphan -eq 'Yes' | Out-GridView -Title 'Select orphaned profiles to delete (Ctrl/Cmd-click to multi-select)' -PassThru
    if (-not $selected) { Write-Host 'No selection made.'; return }
    $toRemove = $selected
} else {
    # Non-interactive or console-only: auto-select orphaned profiles and ask for confirmation
    $toRemove = $gridItems | Where-Object Orphan -eq 'Yes'
    if (-not $toRemove) { Write-Host 'No orphaned profiles found.'; return }
    Write-Host "Found $($toRemove.Count) orphaned profile(s)."
    $confirm = Read-Host "Type 'YES' to confirm deletion of these profiles"
    if ($confirm -ne 'YES') { Write-Host 'Aborted.'; return }
}

if (-not $toRemove -or $toRemove.Count -eq 0) { Write-Host 'No profiles selected for deletion.'; return }

foreach ($it in $toRemove) {
    $path = $it.Path
    if ($PSCmdlet.ShouldProcess($path, 'Remove user profile')) {
        try {
            Remove-CimInstance -InputObject $it.ProfileObj -ErrorAction Stop
            Write-Verbose "Deleted profile $path"
        } catch {
            $err = $_.Exception.Message
            Write-Warning ("Failed to delete profile {0}: {1}" -f $path, $err)
        }
    } else {
        Write-Verbose "Skipping deletion of $path"
    }
}

Write-Host 'Operation complete.' -ForegroundColor Green

} catch {
    # Log unhandled exceptions so ps2exe builds produce a record
    try { Write-Log $_.Exception.ToString() } catch {}
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        $msg = "An unexpected error occurred:\n$($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($msg,'Error',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } catch {}
    exit 1
}

