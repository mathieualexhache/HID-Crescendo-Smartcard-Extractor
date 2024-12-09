# Check if the script is running with elevated privileges
function Test-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# If not running as administrator, restart with elevated privileges
if (-not (Test-Admin)) {
    $arguments = "-File `"$($MyInvocation.MyCommand.Path)`""
    Write-Host "Not running as administrator. Restarting with elevated privileges..."
    Start-Process powershell -ArgumentList $arguments -Verb runAs
    exit
}

Write-Host "Running as administrator, proceeding with registry changes..."

# Define the Vendor ID and Product ID for Omnikey HID 3021
$readerVID = "VID_076B"
$readerPID = "PID_3031"

# Path to the parent registry key for all connected Omnikey devices
$baseRegistryPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\USB\$readerVID&$readerPID"

# Get all subkeys (representing different devices/serial numbers) under the USB VID & PID path
$usbDevices = Get-ChildItem -Path $baseRegistryPath -Recurse

# Loop through each device (subkey) that matches
foreach ($device in $usbDevices) {
    # Construct the path to the WUDFUsbccidDriver registry key for each device
    $escapeCommandKeyPath = "$($device.PSPath)\Device Parameters\WUDFUsbccidDriver"

    # Check if the WUDFUsbccidDriver registry key exists for the current device
    if (Test-Path $escapeCommandKeyPath) {
        # Set the EscapeCommandEnable key to 1 to enable escape commands
        Set-ItemProperty -Path $escapeCommandKeyPath -Name "EscapeCommandEnable" -Value 1
        Write-Host "EscapeCommandEnable set to 1 for: $escapeCommandKeyPath"
    } else {
        Write-Host "No WUDFUsbccidDriver path found for device: $($device.PSPath). No changes made."
    }
}

Write-Host "Script execution completed."
exit