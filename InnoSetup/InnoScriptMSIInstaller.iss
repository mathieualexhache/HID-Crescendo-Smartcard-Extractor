; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "HID Smartcard Serial Extractor"
#define MyAppVersion "1.2"
#define MyAppPublisher "Employment and Social Development Canada (ESDC/EDSC)"
#define MyAppExeName "HID Smart Card SerialExtractor.exe"
#define MyAppAssocName MyAppName + " File"
#define MyAppAssocExt ".exe"
#define MyAppAssocKey StringChange(MyAppAssocName, " ", "") + MyAppAssocExt

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{058F8BFB-A126-4941-B217-2FD084AEA05E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\{#MyAppName}
DisableDirPage=yes
; "ArchitecturesAllowed=x64compatible" specifies that Setup cannot run
; on anything but x64 and Windows 11 on Arm.
ArchitecturesAllowed=x64compatible
; "ArchitecturesInstallIn64BitMode=x64compatible" requests that the
; install be done in "64-bit mode" on x64 or Windows 11 on Arm,
; meaning it should use the native 64-bit Program Files directory and
; the 64-bit view of the registry.
ArchitecturesInstallIn64BitMode=x64compatible
ChangesAssociations=yes
DisableProgramGroupPage=yes
; Remove the following line to run in administrative install mode (install for all users.)
PrivilegesRequired=lowest
OutputDir=C:\Temp
OutputBaseFilename=Windows Build Testing
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Users\mathieualex.hache\source\repos\HID Smart Card SerialExtractor\Release\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\mathieualex.hache\source\repos\HID Smart Card SerialExtractor\Release\CheckAndInstallPwsh.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\mathieualex.hache\source\repos\HID Smart Card SerialExtractor\Release\SetEscapeCommand.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\mathieualex.hache\source\repos\HID Smart Card SerialExtractor\Release\SmartCardInventoryUploader.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\mathieualex.hache\source\repos\HID Smart Card SerialExtractor\Release\certutilExpirationDateSmartCard.ps1"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Registry]
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocExt}\OpenWithProgids"; ValueType: string; ValueName: "{#MyAppAssocKey}"; ValueData: ""; Flags: uninsdeletevalue
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}"; ValueType: string; ValueName: ""; ValueData: "{#MyAppAssocName}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#MyAppExeName},0"
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".myp"; ValueData: ""

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "powershell.exe"; Parameters: "-File ""{app}\CheckAndInstallPwsh.ps1"""; Flags: runhidden; StatusMsg: "Checking for PowerShell 7 installation... If not found, downloading and installing it."
Filename: "pwsh.exe"; Parameters: "-Command Install-Module -Name PnP.PowerShell -Scope AllUsers -Force"; Flags: runhidden; StatusMsg: "Installing PnP.PowerShell module for all users..."
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\SetEscapeCommand.ps1"""; Flags: runhidden runascurrentuser waituntilterminated; StatusMsg: "Checking and enabling escape CCID commands for connected Omnikey HID devices..."
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent; StatusMsg: "Launching {#MyAppName}..."