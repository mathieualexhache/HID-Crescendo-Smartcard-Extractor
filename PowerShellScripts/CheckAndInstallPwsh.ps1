if (-not (Get-Command pwsh -ErrorAction SilentlyContinue)) {
    Write-Host 'PowerShell 7 is not installed. Downloading and installing PowerShell 7...'
    
    # Define the URL for the PowerShell 7 installer
    $installerUrl = "https://github.com/PowerShell/PowerShell/releases/download/v7.2.5/PowerShell-7.2.5-win-x64.msi"
    
    # Define the path to save the installer
    $installerPath = "$env:TEMP\PowerShell-7.2.5-win-x64.msi"
    
    # Download the installer
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
    
    # Install PowerShell 7
    Start-Process msiexec.exe -ArgumentList "/i $installerPath /quiet /norestart" -Wait
    
    # Verify installation
    if (Get-Command pwsh -ErrorAction SilentlyContinue) {
        Write-Host 'PowerShell 7 has been installed successfully.'
    } else {
        Write-Host 'Failed to install PowerShell 7. Please install it manually.'
        exit 1
    }
} else {
    Write-Host 'PowerShell 7 is already installed.'
}