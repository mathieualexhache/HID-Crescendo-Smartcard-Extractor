# HID-Crescendo-Smartcard-Extractor

The **HID Smartcard Serial Extractor** is a Windows-based C++ program designed to facilitate batch swiping and serial number extraction from ISO/IEC 7816 contact smart cards. It works with the **HID Crescendo 144K FIPS Contact Only (SCE7.0 144KDI)** smart card model and the **OMNIKEY® 3021 Contact Smart Card Reader**, providing a seamless process for extracting serial numbers and automatically uploading them to a SharePoint Online Inventory Asset Management list. This tool is primarily used for tracking unassigned and assigned smart cards in a networked environment.

## Key Features

- **CCID Support**:  
  The **OMNIKEY® 3021** card reader works seamlessly with Windows systems, eliminating manual driver installation. The program utilizes the CCID specification for communication with the card reader.

- **PC/SC Compliant**:  
  Fully compliant with the **PC/SC framework**, an industry-standard API for smart card readers, enabling interaction with cards via APDU and Extended APDU protocols (for T=1 cards).

- **CPLC Data Extraction**:  
  Uses ISO 7816 commands to extract **Card Production Lifecycle Data (CPLC)**, including serial numbers, from blank smart cards that have not yet been configured with a **PIV applet**. This allows identification of cards before they are provisioned through Credential Management Systems.

- **PIV Data Extraction**:  
  Supports the extraction of **FIPS 201 'Printed Information'** from PIV cards by sending binary APDU commands.

- **Installer Package**:  
  Distributed as a self-contained Windows installer that includes all necessary dependencies, such as PowerShell 7.x and the **PnP.PowerShell** module for SharePoint Online interaction.

- **CSV Report Generation**:  
  Automatically generates a local CSV report containing the extracted serial numbers for inventory tracking.

## System Requirements

- **Operating System**: Windows 10 or later (64-bit only).
- **PowerShell**: PowerShell 7.x or later.
- **Administrator privileges** are required to modify registry settings during installation (specifically, to enable smart card reader escape commands).
- In order to send or receive **Escape commands** to a smart card reader using Microsoft’s CCID driver, you need to modify the Windows registry by adding a DWORD value called `EscapeCommandEnable` and setting it to a non-zero value. This is required to ensure the smart card reader can handle Escape commands. The registry key for enabling Escape CCID commands is located at:  
  `HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\USB\VID_076B&PID_502A\<serial_number>\Device Parameters\WUDFUsbccidDriver`.

## Installation Instructions

This repository includes all the necessary components to build and use the **HID Smartcard Serial Extractor**. The source code, configuration files, and related dependencies are provided, and users can modify and recompile them based on their specific needs.

### Components Included in the Repository:

1. **C++ Source Code**:  
   The core functionality of the program is written in C++ and can be found in the `src/` directory. The file `HID Smart Card SerialExtractor.cpp` contains the logic for interacting with the smart card reader, extracting serial numbers, and invoking the **PnP.PowerShell** module to upload new items to SharePoint Online.

2. **Inno Setup Compiler File**:  
   The installer configuration file (`InnoSetup/SmartCardInstaller.iss`) is included for users who need to generate a custom Windows installer. You can modify this file to adjust installation paths, customize the installation process, or add other preferences.

3. **PowerShell Scripts**:  
   Several PowerShell scripts are used for operations such as installing dependencies, modifying system registry settings, and uploading data to SharePoint Online. These scripts are invoked by the C++ program to ensure the environment is correctly configured and to upload the extracted card data. The relevant PowerShell scripts can be found in the `PowerShellScripts/` directory:
   - `CheckAndInstallPwsh.ps1`: Ensures that **PowerShell 7.x** and the **PnP.PowerShell** module are installed.
   - `SetEscapeCommand.ps1`: Modifies the registry to enable Escape CCID commands, which are necessary for communication with the smart card reader.
   - `SmartCardInventoryUploader.ps1`: Handles the upload of card serial numbers to SharePoint Online.
   - `certutilExpirationDateSmartCard.ps1`: Extracts additional information from smart cards using the **certutil** utility. This script parses **X.509 certificates** stored on the smart card, extracting data such as expiration dates, issuer information, and other certificate details. This information can be useful for card validation and lifecycle management.

### Customization and Recompilation:

If you need to adjust settings, modify functionality, or add new features, follow these steps:

1. **Modify the Source Code**:  
   The C++ source code is fully available, and you can modify it as needed for your specific use case. After making changes, you can recompile the code using an IDE like **Visual Studio** or any compatible C++ build system.

2. **Customize the PowerShell Scripts**:  
   If your environment requires custom actions (e.g., custom SharePoint Online lists, different registry configurations, or additional card-handling logic), you can modify the PowerShell scripts accordingly.

3. **Rebuild the Installer**:  
   After adjusting the source code and PowerShell scripts, you can rebuild the Windows installer using **Inno Setup**. The installer configuration file (`SmartCardInstaller.iss`) is already provided, but you can customize it further to match your desired installation paths or other preferences.

4. **Recompile the Program**:  
   After making changes to the source code or configuration, recompile the program to generate a new `.exe` file. Make sure all necessary dependencies (such as PowerShell, smart card drivers, etc.) are installed before running the compiled executable.

### Building and Running the Program:

Once the program is compiled, the installation process will follow the default instructions provided earlier.

## Logs

The **HID Smartcard Serial Extractor** generates log files in the following directory:  
`C:\Users\<username>\AppData\Local\HID Smartcard Serial Extractor\`. These log files contain detailed execution logs for the **PowerShell script** that uploads the collected card information to **SharePoint Online**.

The log entries indicate the progress of the serial number upload process, including:
- Initializing and detecting changes in the CSV file.
- Connecting to SharePoint Online.
- Processing individual serial numbers (adding new items or updating existing ones).
- Disconnection from SharePoint after the process is complete.


## Using the Smart Card Serial Extractor

### 1. Scanning Smart Cards

- **Connect the Smart Card Reader**:  
  Ensure the **OMNIKEY® 3021** or compatible smart card reader is connected via USB.

- **Launch the Program**:  
  Open the program by double-clicking the built icon.

- **Scan Cards**:  
  Insert a smart card into the reader. The program will automatically extract its serial number and display a prompt to continue scanning more cards or stop.
