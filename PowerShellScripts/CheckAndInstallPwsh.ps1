# Define the Write-Log function
function Write-Log {
    param (
        [string]$message,
        [string]$logFile
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logMessage = "$timestamp - $message"
    
    # Log the message to the console as well
    Write-Host $logMessage
    
    # Append the log message to the log file
    Add-Content -Path $logFile -Value $logMessage
}

function Install-PowerShell7 {
    param (
        [string]$installerUrl = "https://github.com/PowerShell/PowerShell/releases/download/v7.2.5/PowerShell-7.2.5-win-x64.msi",
        [string]$installerPath = "C:\Temp\PowerShell-7.2.5-win-x64.msi",
        [string]$logFile = "C:\Temp\PowerShell7Install.log"
    )

    # Function to check if the script is running with Administrator privileges
    function Test-IsAdmin {
        [OutputType([bool])]
        param ()
        return (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
    }

    Write-Log "Starting PowerShell 7 installation script." $logFile

    # Check if the script is running with admin privileges
    if (-not (Test-IsAdmin)) {
        Write-Log "Script is not running as Administrator. Restarting with elevated privileges..." $logFile
        
        try {
            # Attempt to restart the script as an administrator
            Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
            exit
        } catch {
            Write-Log "Failed to restart the script with elevated privileges. The operation was canceled or blocked." $logFile
            exit 1
        }
    }

    try {
        Write-Log 'Downloading PowerShell 7 installer...' $logFile
        $attempts = 3
        while ($attempts -gt 0) {
            try {
                Write-Log "Attempting to download installer from $installerUrl..." $logFile
                Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -ErrorAction Stop
                Write-Log 'Download completed successfully.' $logFile
                break
            } catch {
                Write-Log "Download failed. Retrying... ($($attempts - 1) attempts left)" $logFile
                $attempts--
                if ($attempts -eq 0) {
                    Write-Log "Failed to download PowerShell 7 installer after multiple attempts." $logFile
                    throw "Failed to download PowerShell 7 installer after multiple attempts."
                }
            }
        }

        Write-Log "Installing PowerShell 7..." $logFile
        # Execute the installer with the necessary arguments
        msiexec.exe /package $installerPath /quiet /log $env:TEMP\pwsh_install.log ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1 USE_MU=1 ENABLE_MU=1 ADD_PATH=1
        
        # Add a short delay to allow the installation to complete fully
        Start-Sleep -Seconds 10  # Wait for 10 seconds

        # Verify installation by checking the Program Files directory
        $pwshPath = "C:\Program Files\PowerShell\7\pwsh.exe"
        if (Test-Path $pwshPath) {When i test to run my powershell script called './SetEscapeCommand.ps1' 
            Write-Log 'PowerShell 7 has been installed successfully.' $logFile
        } else {
            Write-Log 'Failed to install PowerShell 7. Please install it manually.' $logFile
            exit 1
        }

    } catch {
        Write-Log "An error occurred: $_" $logFile
        exit 1
    } finally {
        Write-Log 'Installer file will remain in C:\Temp.' $logFile
    }
}

# Main script execution
$logFile = "C:\Temp\PowerShell7Install.log"

if (-not (Get-Command pwsh -ErrorAction SilentlyContinue)) {
    Write-Log 'PowerShell 7 is not installed. Downloading and installing PowerShell 7...' $logFile
    Install-PowerShell7 -logFile $logFile
} else {
    Write-Log 'PowerShell 7 is already installed.' $logFile
}

# Now that PowerShell 7 is installed (or already present), let's install the PnP.PowerShell module
try {
    # Check if PnP.PowerShell module is already installed using pwsh.exe
    $pnPModuleInstalled = & "C:\Program Files\PowerShell\7\pwsh.exe" -Command "Get-InstalledModule -Name PnP.PowerShell -ErrorAction SilentlyContinue"
    
    if ($pnPModuleInstalled) {
        Write-Log 'PnP.PowerShell module is already installed.' $logFile
    } else {
        Write-Log 'PnP.PowerShell module is not installed. Installing now using pwsh.exe...' $logFile
        # Install PnP.PowerShell module for current user using pwsh.exe
        & "C:\Program Files\PowerShell\7\pwsh.exe" -Command "Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
        Write-Log 'PnP.PowerShell module has been installed successfully.' $logFile
    }
} catch {
    Write-Log "An error occurred while installing the PnP.PowerShell module: $_" $logFile
    exit 1
}