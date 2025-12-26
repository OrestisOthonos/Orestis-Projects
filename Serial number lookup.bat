Write-Host "Computer Name: $env:COMPUTERNAME"

Get-CimInstance -ClassName Win32_BIOS |
    Format-List Manufacturer, Name, SerialNumber

Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID | Select-Object `
    @{Name='Manufacturer';Expression={[System.Text.Encoding]::ASCII.GetString(($_.ManufacturerName | Where-Object {$_ -ne 0}))}}, `
    @{Name='Model';Expression={[System.Text.Encoding]::ASCII.GetString(($_.UserFriendlyName | Where-Object {$_ -ne 0}))}}, `
    @{Name='SerialNumber';Expression={[System.Text.Encoding]::ASCII.GetString(($_.SerialNumberID | Where-Object {$_ -ne 0}))}}

Get-NetIPAddress -InterfaceAlias '*Ethernet*','*Wi*Fi*' -AddressFamily IPv4 |
    Select-Object InterfaceAlias, IPAddress | Format-Table -AutoSize

Get-NetAdapter -Name '*Ethernet*','*Wi*Fi*' |
    Select-Object Name, MacAddress | Format-Table -AutoSize

# BitLocker Recovery Password Protector Info
Get-BitLockerVolume -MountPoint 'C:' |
    Select-Object -ExpandProperty KeyProtector |
    Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }

Read-Host 'Press Enter to exit'
