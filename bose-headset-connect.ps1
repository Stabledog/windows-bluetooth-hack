# "Disable-PnpDevice" and "Enable-PnpDevice" commands require admin rights
#Requires -RunAsAdministrator
# Original source: https://stackoverflow.com/a/71539568/237059

# Substitute it with the name of your audio device.
# The audio device you are trying to connect to should be paired.
$headphonesName = "LE-Stablecat"

$bluetoothDevices = Get-PnpDevice -class Bluetooth

# My headphones are recognized as 3 devices:
# * PX7 Bowers & Wilkins
# * PX7 Bowers & Wilkins Avrcp Transport
# * PX7 Bowers & Wilkins Avrcp Transport
# *
# It is not enough to toggle only 1 of them. We need to toggle all of them.
$headphonePnpDevices = $bluetoothDevices | Where-Object { $_.Name.StartsWith("$headphonesName") }

if(!$headphonePnpDevices) {
    Write-Host "Coudn't find any devices related to the '$headphonesName'"
    Write-Host "Whole list of available devices is:"
    $bluetoothDevices
    return
}

Write-Host "The following PnP devices related to the '$headphonesName' headphones found:"
ForEach($d in $headphonePnpDevices) { $d }

Write-Host "`nDisable all these devices"
ForEach($d in $headphonePnpDevices) { Disable-PnpDevice -InstanceId $d.InstanceId -Confirm:$false }

Write-Host "Enable all these devices"
ForEach($d in $headphonePnpDevices) { Enable-PnpDevice -InstanceId $d.InstanceId -Confirm:$false }

Write-Host "The headphones should be connected now."

# After toggling, Windows "Bluetooth devices" dialog shows the state of the headphones as "Connected" but not as "Connected voice, music"
# However, after some time (around 5 seconds), the state changes to "Connected voice, music".
Write-Host "It may take around 10 seconds until the Windows starts showing audio devices related to these headphones."
