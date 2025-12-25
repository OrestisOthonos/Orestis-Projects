<#
.SYNOPSIS
Deletes orphaned or 'Account Unknown' user profiles from a Windows system.

.DESCRIPTION
This script queries the Win32_UserProfile WMI class. It uses the most reliable method
to identify orphaned profiles by checking if the profile's Security Identifier (SID)
can be translated into an active user account name. If the SID cannot be resolved,
it is considered orphaned.

The script runs in interactive mode, prompting for confirmation for each profile.
The user can type 'A' to approve the deletion of all subsequent orphaned profiles.

.NOTES
Requires administrative privileges to run.
#>


# Requires elevated privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show("This script must be run as Administrator.", "Error", 'OK', 'Error') | Out-Null
    exit 1
}

# --- Function to check if a SID is linked to an active account ---
function Test-SidResolution {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SID
    )
    try {
        $SecurityIdentifier = New-Object System.Security.Principal.SecurityIdentifier $SID
        $Account = $SecurityIdentifier.Translate([System.Security.Principal.NTAccount])
        if ($Account.Value -notlike "*$SID*") {
            return $true
        }
    } catch {
        return $false
    }
    return $false
}
# -----------------------------------------------------------------


# --- Modern WPF Window UI ---
Add-Type -AssemblyName PresentationFramework
$Profiles = Get-CimInstance -ClassName Win32_UserProfile | Where-Object { -not $_.Special -and -not $_.Loaded }

$ProfileList = @()
foreach ($UserProfile in $Profiles) {
    $SID = $UserProfile.SID
    $Path = $UserProfile.LocalPath
    $IsOrphan = -not (Test-SidResolution -SID $SID)
    $ProfileList += [PSCustomObject]@{
        SID = $SID
        Path = $Path
        Orphan = $IsOrphan
        ProfileObj = $UserProfile
    }
}


$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                Title="User Profile Cleanup" Height="450" Width="800" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <DataGrid x:Name="dgProfiles" AutoGenerateColumns="False" SelectionMode="Extended" CanUserAddRows="False" Grid.Row="0" Margin="0,0,0,10">
            <DataGrid.Columns>
                <DataGridCheckBoxColumn Header="Delete" Binding="{Binding Delete}" Width="60"/>
                <DataGridTextColumn Header="Profile Path" Binding="{Binding Path}" Width="*"/>
                <DataGridTextColumn Header="SID" Binding="{Binding SID}" Width="250"/>
                <DataGridTextColumn Header="Orphaned" Binding="{Binding Orphan}" Width="80"/>
            </DataGrid.Columns>
        </DataGrid>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnDelete" Content="Delete Selected" Width="120" Margin="0,0,10,0"/>
            <Button x:Name="btnExit" Content="Exit" Width="80"/>
        </StackPanel>
    </Grid>
</Window>
"@



$stringReader = New-Object System.IO.StringReader $xaml
$xmlReader = [System.Xml.XmlReader]::Create($stringReader)
try {
    $window = [Windows.Markup.XamlReader]::Load($xmlReader)
} catch {
    [System.Windows.MessageBox]::Show("Failed to load window.\n$($_.Exception.Message)", "Error", 'OK', 'Error') | Out-Null
    exit 1
}

# Prepare data for DataGrid
$data = $ProfileList | ForEach-Object {
    [PSCustomObject]@{
        Delete = $_.Orphan
        Path = $_.Path
        SID = $_.SID
        Orphan = if ($_.Orphan) { 'Yes' } else { 'No' }
        ProfileObj = $_.ProfileObj
    }
}

$dg = $window.FindName('dgProfiles')
if ($null -eq $dg) {
    [System.Windows.MessageBox]::Show("Failed to find DataGrid in window.", "Error", 'OK', 'Error') | Out-Null
    exit 1
}
$dg.ItemsSource = $data

# Button handlers
$btnDelete = $window.FindName('btnDelete')
$btnExit = $window.FindName('btnExit')

if ($null -eq $btnDelete -or $null -eq $btnExit) {
    [System.Windows.MessageBox]::Show("Failed to find buttons in window.", "Error", 'OK', 'Error') | Out-Null
    exit 1
}

$btnDelete.Add_Click({
    $toDelete = @($dg.ItemsSource | Where-Object { $_.Delete })
    if ($toDelete.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No profiles selected for deletion.", "Info", 'OK', 'Information') | Out-Null
        return
    }
    $msg = "Are you sure you want to delete the selected profiles? This cannot be undone."
    $result = [System.Windows.MessageBox]::Show($msg, "Confirm Deletion", 'YesNo', 'Warning')
    if ($result -eq 'Yes') {
        foreach ($item in $toDelete) {
            try {
                Remove-CimInstance -InputObject $item.ProfileObj
                $data.Remove($item)
            } catch {
                [System.Windows.MessageBox]::Show("Failed to delete profile: $($item.Path)\n$($_.Exception.Message)", "Error", 'OK', 'Error') | Out-Null
            }
        }
        $dg.Items.Refresh()
        [System.Windows.MessageBox]::Show("Selected profiles deleted.", "Done", 'OK', 'Information') | Out-Null
    }
})

$btnExit.Add_Click({ $window.Close() })

[void]$window.ShowDialog()

# SIG # Begin signature block
# MIIe6QYJKoZIhvcNAQcCoIIe2jCCHtYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBYuBIHvPeOLYWY
# qM94v5zVK7dNWXtzvNL60HAUELu2GKCCA2AwggNcMIICRKADAgECAhBg0d5ZM/rs
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
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIC2L
# x8do0zF84m2z5iFNvlobBwabxlYqXGXfW+lOsovGMA0GCSqGSIb3DQEBAQUABIIB
# AIP9ujTTj1d0TY1MwIfUxTnsT5wR6hjKzw144qtLm/5fiOqOZm5xuQKRv9uuk4Sn
# WNmDWicUcoNbkfYsUcdjwYI/PQYvS/p4OP/Aqs+XaD1hqyULNSu2vi5L+gdpTkP8
# eNnWHePN7MzZ/XYT2iEEThfmS9wDAaEz+W6MB1cLV+Mq3i7MRw55oEuRwdHfnbG2
# aEu8m14LDiKyAvEUAPKKkSiklEJD3Zk0KWpWQ1NNcgw84c9oEv3r6rnae/Nj4+xR
# NdJxE1rNisuhFqHCWIS9ob1xSW7E98rbKTPu5YNk+6RlD3XeqgGOIoDc9GUNh1ll
# Aa7fNvvlFUcUSmoLIurH8M+hghjYMIIY1AYKKwYBBAGCNwMDATGCGMQwghjABgkq
# hkiG9w0BBwKgghixMIIYrQIBAzEPMA0GCWCGSAFlAwQCAgUAMIH4BgsqhkiG9w0B
# CRABBKCB6ASB5TCB4gIBAQYKKwYBBAGyMQIBATAxMA0GCWCGSAFlAwQCAQUABCBf
# ld77ZnxRxRhNk/xqIMX87vG4JDMVbpeudI9zvkrPmAIVALDAFOQJTUgFOFpwlJ0E
# iyfYumGfGA8yMDI1MTIyMzE2MjY1MlqgdqR0MHIxCzAJBgNVBAYTAkdCMRcwFQYD
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
# CSqGSIb3DQEJBTEPFw0yNTEyMjMxNjI2NTJaMD8GCSqGSIb3DQEJBDEyBDBoj7Pg
# 1earl4x5qiMq252ljj4fYiN4uHbZzjUeZsnhnaVj7HcOPF/EB3YIQWrGV5wwggF6
# BgsqhkiG9w0BCRACDDGCAWkwggFlMIIBYTAWBBQ4yRSBEES03GY+k9R0S4FBhqm1
# sTCBhwQUxq5U5HiG8Xw9VRJIjGnDSnr5wt0wbzBbpFkwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NgIQeiOu2lNplg+RyD5c9MfjPzCBvAQU
# hT1jLZOCgmF80JA1xJHeksFC2scwgaMwgY6kgYswgYgxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkgQ2l0eTEeMBwGA1UE
# ChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVVU0VSVHJ1c3QgUlNB
# IENlcnRpZmljYXRpb24gQXV0aG9yaXR5AhA2wrC9fBs656Oz3TbLyXVoMA0GCSqG
# SIb3DQEBAQUABIICAAiG73+97ikZfLXWeHuswu/MbYnSJrJzJjc4ckR38boYWMBh
# k3NrRUvgRzMqbBUuz02bA9dz3YSCZE8ZyYUDDHS2EM+S1iQwLDrlAwGYqM+TNExN
# CH3s18eFFktDgRO8t6/86kk/MgG0k8V53lbSnoBFmGvWjSKM7emZrJEoBsqqJxer
# aKUHAh85eor6UnDQn4e2oCszUwaE+SpIqXtX4IpW3hz3M441x5Sc+yXV+92ucIrX
# i86JD98XP1ON40Vq3iRG1F2tMqA/+EA6fTJsYbZKaBLrj1pSB6l8dMJILthUGkDz
# QnX1xCnxB7fAIGF1olm4iPGaghM7va7alZgE/ozP+kzP3PwkgbLo1ggRM8GVCLiT
# 2P38YxqDH0QdqcHNecQ2624pYILb5xg+Z/lhWgsKMrE/zJPgpNpY4kVNoH77nByD
# +GoJ15FQXmfcgq7G5W+lY8Dn9Ho4Sp+MT1AGU+nt8wgvPr+JmsXSHLv4GvJRn0D7
# v0AygZ+ps1Y102MEPiB7lcfsxkAjvBU+MNY4RDG6waRtbhm6cV/2bW+PyRF6f6sr
# 0XxP/AKYFPdKWEYqMtrLmreuQcr7cr0B9J8jJ9eHCnUizK8l0SrBz6HRfA4dNOg3
# MMLk1cQjOIvT0/YjUz9r0b/zw1mTy91Ju4opBuBP901+bW+7IepFo/WfYxSn
# SIG # End signature block
