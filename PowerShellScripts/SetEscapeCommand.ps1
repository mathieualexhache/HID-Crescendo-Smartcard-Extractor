# Define log file path
$logFilePath = "C:\Temp\HID_OMNIKEY_Install_Log.txt"
$MSIInstallationlogFilePath = "C:\Temp\HID_OMNIKEY_MSI_Install_Log.txt"

# Function to write log messages to the log file
function Write-Log {
    param (
        [string]$message
    )
    # Write the message with a timestamp to the log file
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$timestamp - $message" | Out-File -Append -FilePath $logFilePath
    Write-Host $message  # Also print to console for immediate feedback
}

# Check if the script is running with elevated privileges
function Test-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Define the URL for the driver ZIP file and local paths
$installerZipUrl = "https://www.hidglobal.com/sites/default/files/drivers/sfw-01556-revc-hid-omnikey-ccid-driver-2.3.4.121.zip"
$installerZipPath = "C:\Temp\HID_OMNIKEY_Driver.zip"
$extractedFolder = "C:\Temp\HID_OMNIKEY_Extracted"

# Flag to track whether the driver has been installed already
$global:driverInstalled = $false

# Check if the script is running as administrator
if (-not (Test-Admin)) {
    # If not elevated, do the initial download and installation of the driver
    if (-not $global:driverInstalled) {
        Write-Log "Downloading HID OMNIKEY CCID Driver ZIP file from $installerZipUrl..."
        try {
            $userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
            # Use curl with proper headers
            & "C:\Windows\System32\curl.exe" -L -o $installerZipPath --header "User-Agent: $userAgent" $installerZipUrl
            Write-Log "Driver ZIP file downloaded successfully."
        } catch {
            Write-Log "Error downloading the HID OMNIKEY CCID Driver ZIP file: $_"
            exit
        }

        # Extract the ZIP file (if it's not already extracted)
        if (Test-Path $installerZipPath) {
            Write-Log "Extracting HID OMNIKEY CCID Driver from ZIP file..."
            if (-not (Test-Path $extractedFolder)) {
                New-Item -Path $extractedFolder -ItemType Directory
            }
            try {
                Expand-Archive -Path $installerZipPath -DestinationPath $extractedFolder -Force
                Write-Log "Driver extracted successfully."
            } catch {
                Write-Log "Error extracting the ZIP file: $_"
                exit
            }
            # Get the path to the subdirectory containing the MSI file
            $msiFolderPath = "$extractedFolder\sfw-01556-revC-hid-omnikey-ccid-driver-2.3.4.121"
            # Run the MSI installer from the subdirectory
            $msiInstallerPath = "$msiFolderPath\HID-OMNIKEY-CCID-Driver-Installer-2.3.4.121-x64.msi"
            Write-Log "msiInstallerPath: $msiInstallerPath"
            if (Test-Path $msiInstallerPath) {
                Write-Log "Running the HID OMNIKEY CCID Driver MSI installer..."
                Write-Log "Full path to MSI installer: $msiInstallerPath"
                try {
                    # Define the arguments for msiexec
                    $MSIArguments = @(
                        "/i"                          # MSI install command
                        ('"{0}"' -f $msiInstallerPath) # Path to the MSI file
                        "/qf"                          # Quiet installation with full UI
                        "/L*V"                         # Log all installation actions
                        ('"{0}"' -f $MSIInstallationlogFilePath)  # Path to the log file
                        "/norestart"
                    )

                    # Run msiexec with elevated privileges and wait until the process finishes
                    $InsightInstall = Start-Process -FilePath "msiexec.exe" -ArgumentList $MSIArguments -Wait -PassThru -WindowStyle Hidden

                    # Check the exit code of the installer
                    if ($InsightInstall.ExitCode -eq 0) {
                        Write-Log "HID OMNIKEY driver installation completed successfully."
                        # Mark as installed to avoid re-running the installation
                        $global:driverInstalled = $true
                    } else {
                        Write-Log "HID OMNIKEY driver installation failed with ExitCode $($InsightInstall.ExitCode)."
                        exit
                    }
                } catch {
                    Write-Log "Error installing the MSI: $_"
                    exit
                }
            } else {
                Write-Log "MSI installer not found in extracted folder. Exiting..."
                exit
            }
        } else {
            Write-Log "Installer not found in C:\Temp. Exiting..."
            exit
        }
    }

    # After installing the driver, elevate the script to admin level
    Write-Log "Not running as administrator. Restarting with elevated privileges..."
    $arguments = "-NoProfile -File `"$($MyInvocation.MyCommand.Path)`""  # Ensure no profiles are loaded
    Start-Process powershell -ArgumentList $arguments -Verb runAs
    exit
}

# Elevated section: perform admin-specific tasks like querying drivers, modifying registry keys, etc.
Write-Log "Running as administrator, proceeding with further tasks..."

# Start capturing all output (standard output and errors) into the log file after elevation
Start-Transcript -Path $logFilePath -Append

# Define the Vendor ID and Product ID for Omnikey HID 3021
$readerVID = "VID_076B"
$readerPID = "PID_3031"

# Path to the parent registry key for all connected Omnikey devices
$baseRegistryPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\USB\$readerVID&$readerPID"

Write-Log "Verifying HID OMNIKEY driver installation..."

# Use Get-WindowsDriver to check for the HID Global driver
$driverVerification = Get-WindowsDriver -online | Where-Object { $_.ProviderName -eq "HID Global" }

if ($driverVerification) {
    Write-Log "Driver installed successfully. Found drivers:"
    $driverVerification.Driver | ForEach-Object { Write-Log $_ }
} else {
    Write-Log "HID Global driver not found. Installation may have failed."
    exit
}

# Uninstall the Microsoft USB CCID Smartcard Reader (if installed)
Write-Log "Checking for existing Microsoft USB CCID Smartcard Reader driver..."
$existingCCIDDriver = Get-WmiObject Win32_PnPSignedDriver | Where-Object { $_.DeviceName -like "*Usbccid Smartcard Reader*" }

if ($existingCCIDDriver) {
    Write-Log "Uninstalling Microsoft USB CCID Smartcard Reader driver..."
    $deviceID = $existingCCIDDriver.InfName
    try {
        pnputil /delete-driver $deviceID /uninstall /force
        Write-Log "Uninstallation of Microsoft USB CCID Smartcard Reader completed."
    } catch {
        Write-Log "Error uninstalling Microsoft USB CCID Smartcard Reader driver: $_"
    }
} else {
    Write-Log "No Microsoft USB CCID Smartcard Reader driver found."
}

# Get all subkeys (representing different devices/serial numbers) under the USB VID & PID path
$usbDevices = Get-ChildItem -Path $baseRegistryPath -Recurse

# Loop through each device (subkey) that matches and set the EscapeCommandEnable registry key
foreach ($device in $usbDevices) {
    # Construct the path to the WUDFUsbccidDriver registry key for each device
    $escapeCommandKeyPath = "$($device.PSPath)\Device Parameters\WUDFUsbccidDriver"
    
    # Check if the WUDFUsbccidDriver registry key exists for the current device
    if (Test-Path $escapeCommandKeyPath) {
        # Set the EscapeCommandEnable key to 1 to enable escape commands
        try {
            Set-ItemProperty -Path $escapeCommandKeyPath -Name "EscapeCommandEnable" -Value 1 -Type DWord
            Write-Log "EscapeCommandEnable set to 1 for: $escapeCommandKeyPath"
        } catch {
            Write-Log "Error setting EscapeCommandEnable for device: $($device.PSPath). Error: $_"
        }
    } else {
        Write-Log "No WUDFUsbccidDriver path found for device: $($device.PSPath)"
    }
}

Write-Log "Script execution completed."

# Stop capturing all output (standard output and errors) into the log file
Stop-Transcript