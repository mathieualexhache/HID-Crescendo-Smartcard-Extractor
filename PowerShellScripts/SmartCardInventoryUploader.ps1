$siteUrl = "https://014gc.sharepoint.com/sites/GCSIProject"
$listName = "SmartCardInventory"
$logListName = "Card Lifecycle Logs" # Secondary list for lifecycle logs
$scriptDir = "C:\Users\mathieualex.hache\AppData\Local\GCSI HID Smartcard Serial Extractor" #$PSScriptRoot  # Directory where the script is located
$csvFilePath = Join-Path -Path $scriptDir -ChildPath "CardReaderData.csv"  # Path to the CSV file
$logFile = Join-Path -Path $scriptDir -ChildPath "smart_card_upload_log.txt"  # Log file to keep track of execution history
$lastRunFile = Join-Path -Path $scriptDir -ChildPath "last_run_time.txt"  # File to store the last successful execution timestamp
$lastFileHashFile = Join-Path -Path $scriptDir -ChildPath "last_file_hash.txt"  # File to store the last file hash

# Set console output encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Function to log messages to the log file
function Log-Message($message) {
    Add-Content -Path $logFile -Value "$(Get-Date) - $message"
}

# Check if the last run time file exists, if not, initialize it and proceed with execution
if (-not (Test-Path $lastRunFile)) {
    Write-Output "[INFO] No previous run time found, initializing."
    Set-Content -Path $lastRunFile -Value (Get-Date)
    Log-Message "[INFO] No previous run time found, initializing."
}

# Read the last successful run time
$lastRunTime = Get-Content -Path $lastRunFile | Out-String
$lastRunTime = [datetime]::Parse($lastRunTime.Trim())

# Check if the CSV file exists before continuing
if (-not (Test-Path $csvFilePath)) {
    Write-Output "[ERROR] CSV file not found. Exiting script."
    Log-Message "[ERROR] CSV file not found. Exiting script."
    return
}

# Compute the hash of the current CSV file
$csvFileHash = Get-FileHash -Path $csvFilePath -Algorithm SHA256

# Check if the hash file exists and compare the hashes
if (Test-Path $lastFileHashFile) {
    $lastFileHash = Get-Content -Path $lastFileHashFile | Out-String
    $lastFileHash = $lastFileHash.Trim()

    if ($csvFileHash.Hash -eq $lastFileHash) {
        Write-Output "[INFO] No new changes detected in the CSV file. Skipping upload."
        Log-Message "[INFO] No new changes detected in the CSV file. Skipping upload."
        return
    }
}

Write-Output "[INFO] New changes detected in the CSV file. Proceeding with upload."
Log-Message "[INFO] New changes detected in the CSV file. Proceeding with upload."

## Connect to SharePoint Online ##
try {
    $sourceConnection = Connect-PnPOnline -Url $siteUrl -UseWebLogin -ReturnConnection
    Write-Output "[SUCCESS] Connected to $siteUrl"
    Log-Message "[SUCCESS] Connected to $siteUrl"
} catch {
    Write-Error "[ERROR] Could not connect to $siteUrl"
    Log-Message "[ERROR] Could not connect to $siteUrl"
    Log-Message $_
    return
}

# Function to get the current user's email address from Azure AD
function Get-CurrentUserEmail {
    $user = Get-PnPUser -Identity "i:0#.f|membership|$env:USERNAME@hrsdc-rhdcc.gc.ca" -Connection $sourceConnection
    return $user.Email
}

# Get the current user
$currentUserEmail = Get-CurrentUserEmail

# Ensure the current user is found
if (-not $currentUserEmail) {
    Write-Error "[ERROR] Could not retrieve current user (Support Tech) information."
    Log-Message "[ERROR] Could not retrieve current user (Support Tech) information."
    return
}

# Print the retrieved user email on the screen
Write-Output "[INFO] Support Tech recorded: $currentUserEmail"
Log-Message "[INFO] Support Tech recorded: $currentUserEmail"

## Read CSV File ##
try {
    $cardData = Import-Csv -Path $csvFilePath -Encoding UTF8
    Write-Output "[SUCCESS] Read the CSV file."
    Log-Message "[SUCCESS] Read the CSV file."
} catch {
    Write-Error "[ERROR] Failed to read the CSV file from $csvFilePath"
    Log-Message "[ERROR] Failed to read the CSV file from $csvFilePath"
    Log-Message $_
    return
}

## Add or Update items in SharePoint Online List ##
foreach ($card in $cardData) {
    try {
        # Ensure Serial Number is present
        if (-not $card.SerialNumber) {
            Write-Warning "[INFO] Skipping entry as Serial Number is missing."
            Log-Message "[INFO] Skipping entry as Serial Number is missing."
            continue
        }

        Write-Output "[INFO] Processing Serial Number: $($card.SerialNumber)"
        Log-Message "[INFO] Processing Serial Number: $($card.SerialNumber)"

        # Query to find if an item with the same Serial Number already exists in the SharePoint list
        $existingItem = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>$($card.SerialNumber)</Value></Eq></Where></Query></View>" -Connection $sourceConnection | Select-Object -First 1

        if ($existingItem) {
            # If the card exists in the SharePoint list and has user information, update it
            Write-Output "[INFO] Card with Serial Number $($card.SerialNumber) already exists in the inventory. Updating details."
            Log-Message "[INFO] Card with Serial Number $($card.SerialNumber) already exists in the inventory. Updating details."

            # Only update if user data (Name, ExpirationDate, ClassifiedEmail) is provided
            if ($card.Name -or $card.ExpirationDate -or $card.ClassifiedEmail) {
                $fieldsToUpdate = @{}

                if ($card.Name) {
                    $fieldsToUpdate["ClientName"] = $card.Name  # Update Name field
                }
                if ($card.ExpirationDate) {
                    $expirationDate = [datetime]::ParseExact($card.ExpirationDate, 'yyyy-MM-dd', $null)
                    $formattedExpirationDate = $expirationDate.ToString('MM/dd/yyyy')
                    $fieldsToUpdate["ExpiryDate"] = $formattedExpirationDate  # Update Expiration Date field
                }
                if ($card.ClassifiedEmail) {
                    $fieldsToUpdate["ClassifiedEmail"] = $card.ClassifiedEmail  # Update Classified Email field
                }

                # If we have any fields to update, proceed with the update
                if ($fieldsToUpdate.Count -gt 0) {
                    Set-PnPListItem -List $listName -Identity $existingItem.Id -Values $fieldsToUpdate -Connection $sourceConnection
                    Write-Output "[INFO] Updated the card with Serial Number $($card.SerialNumber)."
                    Log-Message "[INFO] Updated the card with Serial Number $($card.SerialNumber)."
                }
            }

        } else {
            # If the card does not exist in SharePoint, create a new item
            Write-Output "[INFO] No existing item found for Serial Number $($card.SerialNumber). Adding new item."
            Log-Message "[INFO] No existing item found for Serial Number $($card.SerialNumber). Adding new item."

            # Initialize the new item with Serial Number and default Status
            $item = @{
                "Title" = $card.SerialNumber  # Card Serial Number as Title
                "Status" = 'Unassigned'      # Default status to Unassigned (this will change below if user info is present)
                "ExpiryDate" = $null             # Expiration Date (null if not provided)
                "ClientName" = $null             # Client Name (null if not provided)
                "ClassifiedEmail" = $null     # User Email (null if not provided)
            }

            # If the CSV contains user-related information, populate those fields
            if ($card.Name) {
                $item["ClientName"] = $card.Name
                $item["Status"] = 'Assigned'  # Set status to Assigned when user info is present
            }
            if ($card.ExpirationDate) {
                $expirationDate = [datetime]::ParseExact($card.ExpirationDate, 'yyyy-MM-dd', $null)
                $formattedExpirationDate = $expirationDate.ToString('MM/dd/yyyy')
                $item["ExpiryDate"] = $formattedExpirationDate
            }
            if ($card.ClassifiedEmail) {
                $item["ClassifiedEmail"] = $card.ClassifiedEmail
            }

            # Add the new item to SmartCardInventory
            Add-PnPListItem -List $listName -Values $item -Connection $sourceConnection
            Write-Output "[INFO] Added new item for Serial Number $($card.SerialNumber)."
            Log-Message "[INFO] Added new item for Serial Number $($card.SerialNumber)."
            
            # Get the ID of the card in the SmartCardInventory list (for example, based on SerialNumber)
            $smartCardItem = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>$($card.SerialNumber)</Value></Eq></Where></Query></View>" -Connection $sourceConnection | Select-Object -First 1
            
            # Get current date and time in the correct format (MM/dd/yyyy HH:mm)
            $currentDateTime = (Get-Date).ToString('MM/dd/yyyy HH:mm')
          
            if ($smartCardItem) { 
                $smartCardID = $smartCardItem.Id  # Get the ID of the item
                # Create a log entry in CardLifecycleLogs with the CardNumber Lookup field
                $logItem = @{
                    "CardNumber" = $smartCardID # Set the CardNumber Lookup to the ID of the smart card item
                    "Date" = $currentDateTime # Current datetime in ISO format
                    "Status" = "Placed into Database"  # Status
                    "Information" = "Smart card added to inventory using GCSI smart card batch swiping program."  # Descriptive text
                    "SupportTech" = $currentUserEmail  # Populate the Support Tech field with the current user's email
                }
                
                # Add the log entry to the CardLifecycleLogs list
                Add-PnPListItem -List $logListName -Values $logItem -Connection $sourceConnection
                Write-Output "[INFO] Added new lifecycle log for Serial Number $($card.SerialNumber)."
            } else {
    Write-Error "[ERROR] Smart card with Serial Number $($card.SerialNumber) not found in SmartCardInventory."
}
   }
    } catch {
        Write-Error "[ERROR] Failed to process item for Serial Number: $($card.SerialNumber)"
        Log-Message "[ERROR] Failed to process item for Serial Number: $($card.SerialNumber)"
        Log-Message $_
    }
}

## Update last run time ##
Set-Content -Path $lastRunFile -Value (Get-Date)

# Save the new file hash to detect future changes
Set-Content -Path $lastFileHashFile -Value $csvFileHash.Hash

## Disconnect from SharePoint Online ##
$sourceConnection = $null
Write-Output "[INFO] Disconnected from $siteUrl"
Log-Message "[INFO] Disconnected from $siteUrl"
