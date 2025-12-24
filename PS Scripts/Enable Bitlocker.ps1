<#
.SYNOPSIS
  Enables BitLocker on a drive if not already enabled, otherwise only exports existing recovery keys.
  Saves recovery key(s) to a timestamped file on the current user's Desktop.

.PARAMETER MountPoint
  The drive to act on (default: C:)

.NOTES
  Run in an elevated PowerShell session.
#>
Enable-BitLocker -MountPoint "C" `
                         -EncryptionMethod XtsAes256 `
                         -UsedSpaceOnly `
                         -TpmProtector `
                         -SkipHardwareTest `
                         -ErrorAction Stop

function Write-RecoveryInfo {
    param([Parameter(Mandatory=$true)]$BitLockerVolume)

    $recoveryLines = @()
    foreach ($kp in $BitLockerVolume.KeyProtector) {
        # Support both 'RecoveryPassword' and legacy 'NumericalPassword' type naming
        $isRecoveryType = $kp.KeyProtectorType -in @('RecoveryPassword','NumericalPassword')
        $hasRecoveryProp = $kp.PSObject.Properties.Match('RecoveryPassword').Count -gt 0 -and $kp.RecoveryPassword

        if ($isRecoveryType -or $hasRecoveryProp) {
            $recovery = if ($hasRecoveryProp) { $kp.RecoveryPassword } else { $null }
            $recoveryLines += ("Drive: {0} - Recovery Password: {1} - ID: {2}" -f $BitLockerVolume.MountPoint, $recovery, $kp.KeyProtectorId)
        }
    }

    if ($recoveryLines.Count -gt 0) {
        $recoveryLines | Out-File -FilePath $outputFile -Encoding UTF8 -Force
        Write-Host "Recovery key(s) saved to: $outputFile" -ForegroundColor Green
    } else {
        Write-Warning "BitLocker is enabled on $($BitLockerVolume.MountPoint), but no Recovery Password protector was found. No keys were written."
    }
}

if ($protectionOn) {
    Write-Host "BitLocker is already enabled on $MountPoint. Exporting existing recovery key(s) only..." -ForegroundColor Cyan
    Write-RecoveryInfo -BitLockerVolume $blv
} else {
    Write-Host "BitLocker is not enabled on $MountPoint. Enabling with TPM + Recovery Password..." -ForegroundColor Cyan
    try {
        # Enable BitLocker and add both TPM and Recovery Password protectors
        Enable-BitLocker -MountPoint $MountPoint `
                         -EncryptionMethod XtsAes256 `
                         -UsedSpaceOnly `
                         -TpmProtector `
                         -RecoveryPasswordProtector `
                         -SkipHardwareTest `
                         -ErrorAction Stop

        # Re-query to capture the new recovery password protector
        $blv = Get-BitLockerVolume -MountPoint $MountPoint -ErrorAction Stop
        Write-RecoveryInfo -BitLockerVolume $blv
    } catch {
        Write-Error "Failed to enable BitLocker or export recovery key. $_"
        exit 1
    }
}

# SIG # Begin signature block
# MIIe6AYJKoZIhvcNAQcCoIIe2TCCHtUCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA/7J6lrOI191sK
# oWknZ4KTyKkvNUgXyCD+VUThWSARr6CCA2AwggNcMIICRKADAgECAhBg0d5ZM/rs
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
# qat7ndHgKysInhT+KntTEl3NOZGSt0M37jAxghreMIIa2gIBATBaMEYxFTATBgNV
# BAMMDFNraWxsIE9uIE5ldDEtMCsGCSqGSIb3DQEJARYeb3Jlc3Rpcy5vdGhvbm9z
# QHNraWxsb25uZXQuY29tAhBg0d5ZM/rsvETINKKp+/heMA0GCWCGSAFlAwQCAQUA
# oHwwEAYKKwYBBAGCNwIBDDECMAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIFC6
# vZoKzSWU+RcLwbPIhVzymsn1WhsH8tcZxVwCd+jmMA0GCSqGSIb3DQEBAQUABIIB
# AIzFdyEsa0RAe7E85gjGex7BHrTvJud5vbeyzKNC75Qg/iHFgNb3WGWZxxxu8Nkk
# ErMT4TT9DaVS+Ow+qu7Awsg9ocM/r+ULOK1NU5UMLD1LanaHCpk599atizI/7l9M
# p5ItgNMP68afBxpT7b0DGQdDi1iAEzyo7IfLEcXRdv+a+C98eDCsGCHIy2gAcSCr
# J/PHpwEAtPmeDFC8sWzBvmIBNR+LKt1FvO0vBTVoKPYJ/UWnnbyA0H2R3Y52XHO7
# 6T461G26AhYtfpTZQoX7yWCC0vh5kjPImwm594vNcxSkB+vevEGt4T9NQcolW7bS
# kbafSRAWjIxO8LGuWKHvSMqhghjXMIIY0wYKKwYBBAGCNwMDATGCGMMwghi/Bgkq
# hkiG9w0BBwKgghiwMIIYrAIBAzEPMA0GCWCGSAFlAwQCAgUAMIH3BgsqhkiG9w0B
# CRABBKCB5wSB5DCB4QIBAQYKKwYBBAGyMQIBATAxMA0GCWCGSAFlAwQCAQUABCDO
# SEVnz038R2Ks5BIaKmBaICGRrxcY5uTovVvnKy4zBwIUelyqYd2a1VqPtJKY6D5b
# fwimdbMYDzIwMjUxMjIzMTYyNjU4WqB2pHQwcjELMAkGA1UEBhMCR0IxFzAVBgNV
# BAgTDldlc3QgWW9ya3NoaXJlMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxMDAu
# BgNVBAMTJ1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgU2lnbmVyIFIzNqCC
# EwQwggZiMIIEyqADAgECAhEApCk7bh7d16c0CIetek63JDANBgkqhkiG9w0BAQwF
# ADBVMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSwwKgYD
# VQQDEyNTZWN0aWdvIFB1YmxpYyBUaW1lIFN0YW1waW5nIENBIFIzNjAeFw0yNTAz
# MjcwMDAwMDBaFw0zNjAzMjEyMzU5NTlaMHIxCzAJBgNVBAYTAkdCMRcwFQYDVQQI
# Ew5XZXN0IFlvcmtzaGlyZTEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMTAwLgYD
# VQQDEydTZWN0aWdvIFB1YmxpYyBUaW1lIFN0YW1waW5nIFNpZ25lciBSMzYwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDThJX0bqRTePI9EEt4Egc83JSB
# U2dhrJ+wY7JgReuff5KQNhMuzVytzD+iXazATVPMHZpH/kkiMo1/vlAGFrYN2P7g
# 0Q8oPEcR3h0SftFNYxxMh+bj3ZNbbYjwt8f4DsSHPT+xp9zoFuw0HOMdO3sWeA1+
# F8mhg6uS6BJpPwXQjNSHpVTCgd1gOmKWf12HSfSbnjl3kDm0kP3aIUAhsodBYZsJ
# A1imWqkAVqwcGfvs6pbfs/0GE4BJ2aOnciKNiIV1wDRZAh7rS/O+uTQcb6JVzBVm
# PP63k5xcZNzGo4DOTV+sM1nVrDycWEYS8bSS0lCSeclkTcPjQah9Xs7xbOBoCdma
# hSfg8Km8ffq8PhdoAXYKOI+wlaJj+PbEuwm6rHcm24jhqQfQyYbOUFTKWFe901Vd
# yMC4gRwRAq04FH2VTjBdCkhKts5Py7H73obMGrxN1uGgVyZho4FkqXA8/uk6nkzP
# H9QyHIED3c9CGIJ098hU4Ig2xRjhTbengoncXUeo/cfpKXDeUcAKcuKUYRNdGDlf
# 8WnwbyqUblj4zj1kQZSnZud5EtmjIdPLKce8UhKl5+EEJXQp1Fkc9y5Ivk4AZacG
# MCVG0e+wwGsjcAADRO7Wga89r/jJ56IDK773LdIsL3yANVvJKdeeS6OOEiH6hpq2
# yT+jJ/lHa9zEdqFqMwIDAQABo4IBjjCCAYowHwYDVR0jBBgwFoAUX1jtTDF6omFC
# jVKAurNhlxmiMpswHQYDVR0OBBYEFIhhjKEqN2SBKGChmzHQjP0sAs5PMA4GA1Ud
# DwEB/wQEAwIGwDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMI
# MEoGA1UdIARDMEEwNQYMKwYBBAGyMQECAQMIMCUwIwYIKwYBBQUHAgEWF2h0dHBz
# Oi8vc2VjdGlnby5jb20vQ1BTMAgGBmeBDAEEAjBKBgNVHR8EQzBBMD+gPaA7hjlo
# dHRwOi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1lU3RhbXBpbmdD
# QVIzNi5jcmwwegYIKwYBBQUHAQEEbjBsMEUGCCsGAQUFBzAChjlodHRwOi8vY3J0
# LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1lU3RhbXBpbmdDQVIzNi5jcnQw
# IwYIKwYBBQUHMAGGF2h0dHA6Ly9vY3NwLnNlY3RpZ28uY29tMA0GCSqGSIb3DQEB
# DAUAA4IBgQACgT6khnJRIfllqS49Uorh5ZvMSxNEk4SNsi7qvu+bNdcuknHgXIaZ
# yqcVmhrV3PHcmtQKt0blv/8t8DE4bL0+H0m2tgKElpUeu6wOH02BjCIYM6HLInbN
# HLf6R2qHC1SUsJ02MWNqRNIT6GQL0Xm3LW7E6hDZmR8jlYzhZcDdkdw0cHhXjbOL
# smTeS0SeRJ1WJXEzqt25dbSOaaK7vVmkEVkOHsp16ez49Bc+Ayq/Oh2BAkSTFog4
# 3ldEKgHEDBbCIyba2E8O5lPNan+BQXOLuLMKYS3ikTcp/Qw63dxyDCfgqXYUhxBp
# XnmeSO/WA4NwdwP35lWNhmjIpNVZvhWoxDL+PxDdpph3+M5DroWGTc1ZuDa1iXmO
# FAK4iwTnlWDg3QNRsRa9cnG3FBBpVHnHOEQj4GMkrOHdNDTbonEeGvZ+4nSZXrwC
# W4Wv2qyGDBLlKk3kUW1pIScDCpm/chL6aUbnSsrtbepdtbCLiGanKVR/KC1gsR0t
# C6Q0RfWOI4owggYUMIID/KADAgECAhB6I67aU2mWD5HIPlz0x+M/MA0GCSqGSIb3
# DQEBDAUAMFcxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQx
# LjAsBgNVBAMTJVNlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgUm9vdCBSNDYw
# HhcNMjEwMzIyMDAwMDAwWhcNMzYwMzIxMjM1OTU5WjBVMQswCQYDVQQGEwJHQjEY
# MBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSwwKgYDVQQDEyNTZWN0aWdvIFB1Ymxp
# YyBUaW1lIFN0YW1waW5nIENBIFIzNjCCAaIwDQYJKoZIhvcNAQEBBQADggGPADCC
# AYoCggGBAM2Y2ENBq26CK+z2M34mNOSJjNPvIhKAVD7vJq+MDoGD46IiM+b83+3e
# cLvBhStSVjeYXIjfa3ajoW3cS3ElcJzkyZlBnwDEJuHlzpbN4kMH2qRBVrjrGJgS
# lzzUqcGQBaCxpectRGhhnOSwcjPMI3G0hedv2eNmGiUbD12OeORN0ADzdpsQ4dDi
# 6M4YhoGE9cbY11XxM2AVZn0GiOUC9+XE0wI7CQKfOUfigLDn7i/WeyxZ43XLj5GV
# o7LDBExSLnh+va8WxTlA+uBvq1KO8RSHUQLgzb1gbL9Ihgzxmkdp2ZWNuLc+XyEm
# JNbD2OIIq/fWlwBp6KNL19zpHsODLIsgZ+WZ1AzCs1HEK6VWrxmnKyJJg2Lv23Dl
# EdZlQSGdF+z+Gyn9/CRezKe7WNyxRf4e4bwUtrYE2F5Q+05yDD68clwnweckKtxR
# aF0VzN/w76kOLIaFVhf5sMM/caEZLtOYqYadtn034ykSFaZuIBU9uCSrKRKTPJhW
# vXk4CllgrwIDAQABo4IBXDCCAVgwHwYDVR0jBBgwFoAU9ndq3T/9ARP/FqFsggIv
# 0Ao9FCUwHQYDVR0OBBYEFF9Y7UwxeqJhQo1SgLqzYZcZojKbMA4GA1UdDwEB/wQE
# AwIBhjASBgNVHRMBAf8ECDAGAQH/AgEAMBMGA1UdJQQMMAoGCCsGAQUFBwMIMBEG
# A1UdIAQKMAgwBgYEVR0gADBMBgNVHR8ERTBDMEGgP6A9hjtodHRwOi8vY3JsLnNl
# Y3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1lU3RhbXBpbmdSb290UjQ2LmNybDB8
# BggrBgEFBQcBAQRwMG4wRwYIKwYBBQUHMAKGO2h0dHA6Ly9jcnQuc2VjdGlnby5j
# b20vU2VjdGlnb1B1YmxpY1RpbWVTdGFtcGluZ1Jvb3RSNDYucDdjMCMGCCsGAQUF
# BzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTANBgkqhkiG9w0BAQwFAAOCAgEA
# Etd7IK0ONVgMnoEdJVj9TC1ndK/HYiYh9lVUacahRoZ2W2hfiEOyQExnHk1jkvpI
# JzAMxmEc6ZvIyHI5UkPCbXKspioYMdbOnBWQUn733qMooBfIghpR/klUqNxx6/fD
# XqY0hSU1OSkkSivt51UlmJElUICZYBodzD3M/SFjeCP59anwxs6hwj1mfvzG+b1c
# oYGnqsSz2wSKr+nDO+Db8qNcTbJZRAiSazr7KyUJGo1c+MScGfG5QHV+bps8BX5O
# yv9Ct36Y4Il6ajTqV2ifikkVtB3RNBUgwu/mSiSUice/Jp/q8BMk/gN8+0rNIE+Q
# qU63JoVMCMPY2752LmESsRVVoypJVt8/N3qQ1c6FibbcRabo3azZkcIdWGVSAdoL
# gAIxEKBeNh9AQO1gQrnh1TA8ldXuJzPSuALOz1Ujb0PCyNVkWk7hkhVHfcvBfI8N
# tgWQupiaAeNHe0pWSGH2opXZYKYG4Lbukg7HpNi/KqJhue2Keak6qH9A8CeEOB7E
# ob0Zf+fU+CCQaL0cJqlmnx9HCDxF+3BLbUufrV64EbTI40zqegPZdA+sXCmbcZy6
# okx/SjwsusWRItFA3DE8MORZeFb6BmzBtqKJ7l939bbKBy2jvxcJI98Va95Q5Jnl
# Kor3m0E7xpMeYRriWklUPsetMSf2NvUQa/E5vVyefQIwggaCMIIEaqADAgECAhA2
# wrC9fBs656Oz3TbLyXVoMA0GCSqGSIb3DQEBDAUAMIGIMQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKTmV3IEplcnNleTEUMBIGA1UEBxMLSmVyc2V5IENpdHkxHjAcBgNV
# BAoTFVRoZSBVU0VSVFJVU1QgTmV0d29yazEuMCwGA1UEAxMlVVNFUlRydXN0IFJT
# QSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTAeFw0yMTAzMjIwMDAwMDBaFw0zODAx
# MTgyMzU5NTlaMFcxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLjAsBgNVBAMTJVNlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgUm9vdCBS
# NDYwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCIndi5RWedHd3ouSaB
# mlRUwHxJBZvMWhUP2ZQQRLRBQIF3FJmp1OR2LMgIU14g0JIlL6VXWKmdbmKGRDIL
# RxEtZdQnOh2qmcxGzjqemIk8et8sE6J+N+Gl1cnZocew8eCAawKLu4TRrCoqCAT8
# uRjDeypoGJrruH/drCio28aqIVEn45NZiZQI7YYBex48eL78lQ0BrHeSmqy1uXe9
# xN04aG0pKG9ki+PC6VEfzutu6Q3IcZZfm00r9YAEp/4aeiLhyaKxLuhKKaAdQjRa
# f/h6U13jQEV1JnUTCm511n5avv4N+jSVwd+Wb8UMOs4netapq5Q/yGyiQOgjsP/J
# RUj0MAT9YrcmXcLgsrAimfWY3MzKm1HCxcquinTqbs1Q0d2VMMQyi9cAgMYC9jKc
# +3mW62/yVl4jnDcw6ULJsBkOkrcPLUwqj7poS0T2+2JMzPP+jZ1h90/QpZnBkhdt
# ixMiWDVgh60KmLmzXiqJc6lGwqoUqpq/1HVHm+Pc2B6+wCy/GwCcjw5rmzajLbmq
# GygEgaj/OLoanEWP6Y52Hflef3XLvYnhEY4kSirMQhtberRvaI+5YsD3XVxHGBjl
# Ili5u+NrLedIxsE88WzKXqZjj9Zi5ybJL2WjeXuOTbswB7XjkZbErg7ebeAQUQiS
# /uRGZ58NHs57ZPUfECcgJC+v2wIDAQABo4IBFjCCARIwHwYDVR0jBBgwFoAUU3m/
# WqorSs9UgOHYm8Cd8rIDZsswHQYDVR0OBBYEFPZ3at0//QET/xahbIICL9AKPRQl
# MA4GA1UdDwEB/wQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MBMGA1UdJQQMMAoGCCsG
# AQUFBwMIMBEGA1UdIAQKMAgwBgYEVR0gADBQBgNVHR8ESTBHMEWgQ6BBhj9odHRw
# Oi8vY3JsLnVzZXJ0cnVzdC5jb20vVVNFUlRydXN0UlNBQ2VydGlmaWNhdGlvbkF1
# dGhvcml0eS5jcmwwNQYIKwYBBQUHAQEEKTAnMCUGCCsGAQUFBzABhhlodHRwOi8v
# b2NzcC51c2VydHJ1c3QuY29tMA0GCSqGSIb3DQEBDAUAA4ICAQAOvmVB7WhEuOWh
# xdQRh+S3OyWM637ayBeR7djxQ8SihTnLf2sABFoB0DFR6JfWS0snf6WDG2gtCGfl
# wVvcYXZJJlFfym1Doi+4PfDP8s0cqlDmdfyGOwMtGGzJ4iImyaz3IBae91g50Qyr
# VbrUoT0mUGQHbRcF57olpfHhQEStz5i6hJvVLFV/ueQ21SM99zG4W2tB1ExGL98i
# dX8ChsTwbD/zIExAopoe3l6JrzJtPxj8V9rocAnLP2C8Q5wXVVZcbw4x4ztXLsGz
# qZIiRh5i111TW7HV1AtsQa6vXy633vCAbAOIaKcLAo/IU7sClyZUk62XD0VUnHD+
# YvVNvIGezjM6CRpcWed/ODiptK+evDKPU2K6synimYBaNH49v9Ih24+eYXNtI38b
# yt5kIvh+8aW88WThRpv8lUJKaPn37+YHYafob9Rg7LyTrSYpyZoBmwRWSE4W6iPj
# B7wJjJpH29308ZkpKKdpkiS9WNsf/eeUtvRrtIEiSJHN899L1P4l6zKVsdrUu1FX
# 1T/ubSrsxrYJD+3f3aKg6yxdbugot06YwGXXiy5UUGZvOu3lXlxA+fC13dQ5OlL2
# gIb5lmF6Ii8+CQOYDwXM+yd9dbmocQsHjcRPsccUd5E9FiswEqORvz8g3s+jR3SF
# CgXhN4wz7NgAnOgpCdUo4uDyllU9PzGCBJIwggSOAgEBMGowVTELMAkGA1UEBhMC
# R0IxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQ
# dWJsaWMgVGltZSBTdGFtcGluZyBDQSBSMzYCEQCkKTtuHt3XpzQIh616TrckMA0G
# CWCGSAFlAwQCAgUAoIIB+TAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQwHAYJ
# KoZIhvcNAQkFMQ8XDTI1MTIyMzE2MjY1OFowPwYJKoZIhvcNAQkEMTIEMP/sS/eS
# wcOZcblpcFoPVNDMD17nv9PKhn3eLrns9licwS5VXPW+hJQmhQ/LCtPHtTCCAXoG
# CyqGSIb3DQEJEAIMMYIBaTCCAWUwggFhMBYEFDjJFIEQRLTcZj6T1HRLgUGGqbWx
# MIGHBBTGrlTkeIbxfD1VEkiMacNKevnC3TBvMFukWTBXMQswCQYDVQQGEwJHQjEY
# MBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMS4wLAYDVQQDEyVTZWN0aWdvIFB1Ymxp
# YyBUaW1lIFN0YW1waW5nIFJvb3QgUjQ2AhB6I67aU2mWD5HIPlz0x+M/MIG8BBSF
# PWMtk4KCYXzQkDXEkd6SwULaxzCBozCBjqSBizCBiDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCk5ldyBKZXJzZXkxFDASBgNVBAcTC0plcnNleSBDaXR5MR4wHAYDVQQK
# ExVUaGUgVVNFUlRSVVNUIE5ldHdvcmsxLjAsBgNVBAMTJVVTRVJUcnVzdCBSU0Eg
# Q2VydGlmaWNhdGlvbiBBdXRob3JpdHkCEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZI
# hvcNAQEBBQAEggIAoLjF1uGu5AzWkl/MumAGCwLbUCS3KUvfN76rMCnGaO+r4t7k
# X5OlhhC6+pguC9zaVvkJ/QnswEagM8IjqWRg6zE/PR14tsSQijiaNTimzbBF83Th
# dj3OIUjwn484epCBx5aTuPWtrx3N0QjieUN6yCd1UdKrA4SmshezAcDsWRLtAoEC
# XiR6yvg9kNH0cAHbYV/LKj/4tmLDnZjbH0v0vBhgY1UKELPcmYfC3VYSpSef6zxn
# 91sPMKna3jF35hzs41EiIKtQ59HsKhB6R/SPj74hMfKm3tzzWHiapMvjqv3vYxom
# 6fo/7965enQTUnUg1etB/doP4+yp//LHr/KNd6ExlCf4l/+0zBJh8yXMMHOhJhoq
# XV7W+q8OKSdq47MGUFV2igS5lHmcJxxv8+9LplOLddzKmudfVvYojNH8jkOn6ID9
# 1rSwBNZDRiIcMmKJeoSXXNdW8HTH7PRkpC+XY0eLlNZhG3oHO+RYo7oewp3crm/K
# zhbD9nt7vmdrWdsSitOb0Rq32psIzTXoVeiw0CTrpUN+LvGu8Fh7WrmFnPi2JvOW
# H0aEZL34//89E2Woz5As4mfQFJyaMUVOISFk8yQb9GMjriyUczmw1wNeSHkhwTWE
# YsNg4bjBH5G4Og3mCwJi4JZytgNiPKktWcnmZsqp55f7jTat86H5AoO5c7s=
# SIG # End signature block
