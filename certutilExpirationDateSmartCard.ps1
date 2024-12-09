param(
    [string]$pin  # PIN will be passed as a parameter
)

# Run certutil command and capture output
$certutilOutput = certutil -scinfo -silent -PIN $pin

# Check if the command executed successfully
if ($certutilOutput -eq $null) {
    Write-Error "certutil command failed or produced no output."
    return
}

# Initialize variables to hold the expiration date and email address
$validityDate = $null
$sanEmail = $null

# Search for the first validity period line, which might look something like:
# "NotAfter: 12/31/2024 12:00:00 AM"
$validityPeriodLine = $certutilOutput | Where-Object { $_ -match "NotAfter" } | Select-Object -First 1

# Check if we found the line and extract the date
if ($validityPeriodLine) {
    try {
        # Extract the expiration date from the "NotAfter" line
        # Split the line by colon, then trim the result and extract only the date portion (before the time)
        $validityDate = $validityPeriodLine.Split(":")[1].Trim().Split(" ")[0]
    } catch {
        Write-Error "Error extracting the expiration date from the line."
    }
} else {
    Write-Error "Could not find the validity period in certutil output."
}

# Search for the SubjectAltName line containing the email (e.g., RFC822 Name or Other Name)
$sanLine = $certutilOutput -split "`n" | Where-Object { $_ -match "SubjectAltName" } | Select-Object -First 1

# Check if we found the SAN line and extract the email
if ($sanLine) {
    try {
        # Extract the email address from the SubjectAltName line
        # Corrected regex to match both "Other Name:Principal Name=" and "RFC822 Name="
        if ($sanLine -match 'Other Name:Principal Name=([^\s,]+)|RFC822 Name=([^\s,]+)') {
            # If the match was successful, capture the email address
            $sanEmail = if ($matches[1]) { $matches[1] } else { $matches[2] }
        } else {
            Write-Error "Email pattern not found in SubjectAltName."
        }
    } catch {
        Write-Error "Error extracting the email address from the SubjectAltName."
    }
} else {
    Write-Error "Could not find the SubjectAltName in certutil output."
}

# Combine both the expiration date and email into a single string
if ($validityDate -and $sanEmail) {
    $combinedResult = "$validityDate,$sanEmail"
    Write-Host $combinedResult.Trim()  # Remove any extra newlines or spaces at the end
    return $combinedResult.Trim()      # Ensure no trailing spaces or newlines
} else {
    Write-Error "Could not extract both expiration date and email address."
}