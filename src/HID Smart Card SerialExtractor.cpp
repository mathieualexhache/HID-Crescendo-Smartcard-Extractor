#include <stdio.h>
#include <windows.h>
#include <winscard.h>
#include <wchar.h> // Use MultiByteToWideChar to convert from multi-byte (ASCII) to wide character format.
#include <string>
#include <sstream>
#include <locale>
#include <codecvt> // For converting to UTF-8
#include <cstdio>
#include <iostream>
#include <fstream>
#include <thread>
#include <chrono>
#include <shlwapi.h>
#include <map>
#include <string>
#include <iomanip>
#include <vector>
#include <conio.h>  // For _getch()

using namespace std;

// Global variables for paths
std::wstring exeDir;
std::wstring csvPath;
std::wstring scriptPath;
std::wstring scriptPath2;
int scannedCardCount = 0; // Initialize counter for processed cards

// Function to get the directory of the executable
std::wstring GetExecutableDirectory() {
    wchar_t path[MAX_PATH];
    GetModuleFileName(NULL, path, MAX_PATH);
    std::wstring exePath(path);
    size_t pos = exePath.find_last_of(L"\\/");
    return exePath.substr(0, pos);
}

// Function to check if file is locked
bool IsFileLocked(const std::wstring& filePath) {
    std::ifstream file(filePath);
    return !file.is_open();
}

void MaskedPinInput(char* pin, int minLength, int maxLength) {
    int index = 0;
    char ch;

    std::cout << "\n[PROMPT] Please enter the PIN (6 to 8 digits): ";

    while (index < maxLength) {
        ch = _getch();  // Get a character without echo
        if (ch == 13) {  // Enter key (return) pressed
            break;
        }
        if (ch == 8 && index > 0) {  // Backspace handling
            std::cout << "\b \b";  // Erase the previous character
            index--;
        }
        else if (ch >= '0' && ch <= '9') {  // Only allow digits
            std::cout << '*';  // Display asterisk instead of character
            pin[index++] = ch;
        }
    }

    pin[index] = '\0';  // Null-terminate the string
    std::cout << std::endl;

    // Validate PIN length
    if (index < minLength || index > maxLength) {
        std::cout << "\nInvalid PIN length. The PIN must be between " << minLength << " and " << maxLength << " digits.\n";
        // Clear PIN and prompt again
        memset(pin, 0, maxLength);
        MaskedPinInput(pin, minLength, maxLength);  // Recursively prompt for valid PIN
    }
}

// Function to convert the PIN into the APDU format for verification
void ConvertPinToApdu(char* enteredPIN, BYTE* VERIFY_PIN) {
    int pinLength = strlen(enteredPIN);

    // 5-byte header + max 8 bytes for PIN + filler bytes (0xFF) + 1 byte for control byte (0x00)
    VERIFY_PIN[0] = 0x00;  // CLA
    VERIFY_PIN[1] = 0x20;  // INS (Verify PIN)
    VERIFY_PIN[2] = 0x00;  // P1
    VERIFY_PIN[3] = 0x80;  // P2
    VERIFY_PIN[4] = 0x08;  // Length of PIN field (always 8 bytes, as per APDU format)

    // Convert entered PIN into APDU format (each digit converted to ASCII byte value)
    for (int i = 0; i < pinLength; i++) {
        VERIFY_PIN[5 + i] = enteredPIN[i];  // Store each ASCII character as byte
    }

    // Add the necessary filler bytes (0xFF) for shorter PINs (if the PIN is < 8 digits)
    for (int i = pinLength; i < 8; i++) {
        VERIFY_PIN[5 + i] = 0xFF;  // Fill remaining space with 0xFF
    }

    // Add the trailing 0x00 byte at the end of the APDU command
    VERIFY_PIN[13] = 0x00;
}

std::wstring  ExtractSerialNumber(BYTE* pBuffer, DWORD dwLen) {
    // Check that the response is valid(Status Byte should be 0x9000)
    if (dwLen < 2 || pBuffer[dwLen - 2] != 0x90 || pBuffer[dwLen - 1] != 0x00) {
        std::wcout << L"\n[INFO] Provisioned card detected, UID retrieval failed. The card has already been provisioned through HID ActivID ActivClient." << std::endl;
        return L"";
    }

    // Find the '9F7F' tag and extract the UID starting from the 3rd byte
    DWORD index = 0;
    while (index < dwLen - 2) {
        if (pBuffer[index] == 0x9F && pBuffer[index + 1] == 0x7F) {
            // Length of the data follows after the tag (byte 3)
            DWORD dataLength = pBuffer[index + 2];  // Length of the data

            // Ensure we have enough bytes for the data we're trying to extract
            if (index + 20 > dwLen) {
                std::wcout << L"\n[ERROR] Buffer is too short to extract serial number." << std::endl;
                return L"";
            }

            // Extract the serial number parts into a wstring
            std::wstringstream serialNumber;  // Use wstringstream to build the serial number

            // Segment 1: 00050045 (Bytes 3-10)
            serialNumber << std::uppercase << std::setw(2) << std::setfill(L'0') << std::hex << (int)pBuffer[index + 3]
                << std::setw(2) << std::setfill(L'0') << std::hex << (int)pBuffer[index + 4]
                << std::setw(2) << std::setfill(L'0') << std::hex << (int)pBuffer[index + 5]
                << std::setw(2) << std::setfill(L'0') << std::hex << (int)pBuffer[index + 6];

            // Segment 2: 3AB3 (Bytes 19-20) — This comes after '2214521D'
            serialNumber << std::setw(2) << std::setfill(L'0') << std::hex << (int)pBuffer[index + 19]
                << std::setw(2) << std::setfill(L'0') << std::hex << (int)pBuffer[index + 20];

            // Segment 3: 2214521D (Bytes 15-18) — This comes after the junk section
            serialNumber << std::setw(2) << std::setfill(L'0') << std::hex << (int)pBuffer[index + 15]
                << std::setw(2) << std::setfill(L'0') << std::hex << (int)pBuffer[index + 16]
                << std::setw(2) << std::setfill(L'0') << std::hex << (int)pBuffer[index + 17]
                << std::setw(2) << std::setfill(L'0') << std::hex << (int)pBuffer[index + 18];

            // Return the reconstructed serial number as a wide string
            return serialNumber.str();
        }
        index++;
    }

    return L"";  // If no valid serial number found, return empty string
}

bool RetrieveSerialNumber(SCARDHANDLE hCard, SCARD_IO_REQUEST pIoReq, BYTE* pBuffer, DWORD& dwLen, std::wstring& serialNumber) {
    BYTE APDU_GET_SERIAL[] = { 0x80, 0xCA, 0x9F, 0x7F, 0x00 };
    LONG lResult = SCardTransmit(hCard, &pIoReq, APDU_GET_SERIAL, sizeof(APDU_GET_SERIAL), NULL, pBuffer, &dwLen);
    if (SCARD_S_SUCCESS != lResult) {
        std::wcout << L"\n[ERROR] Failed to retrieve UID. Error code 0x" << std::hex << lResult << std::endl;
        return false;
    }
    // Call ExtractSerialNumber to get the serial number
    serialNumber = ExtractSerialNumber(pBuffer, dwLen);
    return !serialNumber.empty();
}


bool VerifyDefaultPIN(SCARDHANDLE hCard, SCARD_IO_REQUEST pIoReq, BYTE* pBuffer, DWORD& dwLen) {
    BYTE VERIFY_PIN[] = { 0x00, 0x20, 0x00, 0x80, 0x08, 0x31, 0x31, 0x32, 0x32, 0x33, 0x33, 0xFF, 0xFF, 0x00 };
    LONG lResult;
    BYTE GET_PIN_TRY_COUNTER[] = { 0x00, 0x20, 0x00, 0x80, 0x00 };
    printf("\n[INFO] Retrieving the PIN Try Counter...\n");
    dwLen = sizeof(pBuffer);
    lResult = SCardTransmit(hCard, &pIoReq, GET_PIN_TRY_COUNTER, sizeof(GET_PIN_TRY_COUNTER), NULL, pBuffer, &dwLen);
    if (SCARD_S_SUCCESS != lResult) {
        printf("\n[ERROR] Failed to retrieve PIN retry counter. Error code 0x%08X.\n", lResult);
        return false;
    }

    if (dwLen >= 2) {
        BYTE SW1 = pBuffer[dwLen - 2];
        BYTE SW2 = pBuffer[dwLen - 1];
        if (SW1 == 0x63) {
            int retriesLeft = SW2 & 0x0F;
            printf("\n[INFO] Remaining PIN retry tries: %d\n", retriesLeft);
        }
    }

    printf("\n[INFO] Verifying PIN...\n");
    dwLen = sizeof(pBuffer);
    lResult = SCardTransmit(hCard, &pIoReq, VERIFY_PIN, sizeof(VERIFY_PIN), NULL, pBuffer, &dwLen);
    if (SCARD_S_SUCCESS != lResult) {
        printf("\n[ERROR] Default PIN 112233 verification failed. Error code 0x%08X.\n", lResult);
        return false;
    }

    if (dwLen >= 2) {
        BYTE SW1 = pBuffer[dwLen - 2];
        BYTE SW2 = pBuffer[dwLen - 1];
        if (SW1 == 0x90 && SW2 == 0x00) {
            printf("\n[SUCCESS] Default PIN successfully verified. The card is now authenticated.\n");
            return true;
        }
        else {
            printf("\n[ERROR] Default PIN 112233 verification failed. Error code 0x%02X%02X. Please enter your custom PIN.\n", SW1, SW2);
            return false;
        }
    }

    return false;
}

std::string PromptForCustomPIN(SCARDHANDLE hCard, SCARD_IO_REQUEST pIoReq, BYTE* pBuffer, DWORD& dwLen) {
    char enteredPIN[9];  // Assuming the PIN length is between 6 and 8 digits + null terminator
    BYTE VERIFY_PIN[14];  // 5-byte header + max 8 bytes for PIN + filler bytes (0xFF) + 1 byte for control byte (0x00)

    // Step 1: Capture the user-entered PIN with validation
    MaskedPinInput(enteredPIN, 6, 8);  // Enforce PIN length between 6 and 8 digits

    // Step 2: Convert the PIN to APDU format
    ConvertPinToApdu(enteredPIN, VERIFY_PIN);

    // Step 3: Now retry with the new user-provided PIN
    printf("\n[INFO] Verifying custom PIN.\n");

    dwLen = sizeof(pBuffer);  // Ensure pBuffer size is big enough
    LONG lResult = SCardTransmit(hCard, &pIoReq, VERIFY_PIN, sizeof(VERIFY_PIN), NULL, pBuffer, &dwLen);

    if (SCARD_S_SUCCESS != lResult) {
        printf("\n[ERROR] Failed to verify custom PIN. Error code 0x%08X.\n", lResult);
        return ""; // Return an empty string in case of error
    }

    // Check for the SW response indicating success
    USHORT SW = pBuffer[dwLen - 2] << 8 | pBuffer[dwLen - 1];
    if (SW != 0x9000) {
        printf("\n[ERROR] Custom PIN verification failed. Error code 0x%04X.\n", SW);
        return ""; // Return an empty string in case of error
    }

    // Step 4: If custom PIN verification is successful, pass the PIN to the PowerShell script
    printf("\n[INFO] Custom PIN successfully verified. The card is now authenticated.\n");

    return std::string(enteredPIN);
}

void displayScanSummary() {
    // Summary of scanned cards
    printf("\n[SUMMARY] Session complete!\n");
    printf("\n[INFO] Total new cards scanned during this session: %d\n", scannedCardCount);
    // Optionally open a new terminal window displaying the total cards scanned
    // system("start cmd /k echo Total cards scanned: %d && pause", scannedCardCount);
}

void WaitForFileRelease(const std::wstring& filePath) {
    // Wait for the file to be released by C++ before executing the PowerShell script
    while (IsFileLocked(filePath)) {
        std::this_thread::sleep_for(std::chrono::seconds(1));
    }
}

std::pair<std::string, std::string> RunPowerShellCommand(const std::wstring& scriptPath2, const std::string& pin) {
    // Convert the scriptPath (wide string) to a narrow string (std::string)
    std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
    std::string scriptPathStr = converter.to_bytes(scriptPath2);  // Convert wide string to narrow string

    // Build the PowerShell command with the PIN passed as an argument
    std::string psCommand = "powershell.exe -ExecutionPolicy Bypass -File \"" + scriptPathStr + "\" -pin " + pin;

    // Open a pipe to the process
    std::vector<char> buffer(128);  // Use vector instead of std::array
    std::string result = "";
    std::shared_ptr<FILE> pipe(_popen(psCommand.c_str(), "r"), _pclose);

    if (!pipe) {
        throw std::runtime_error("Failed to run PowerShell script.");
    }

    // Read the output of the PowerShell script
    while (fgets(buffer.data(), buffer.size(), pipe.get()) != nullptr) {
        result += buffer.data();
    }

    // Remove any trailing whitespace/newline characters
    result.erase(result.find_last_not_of("\n\r") + 1);

    // Handle the case where no output is returned
    if (result.empty()) {
        throw std::runtime_error("PowerShell script returned empty result.");
    }

    // Now split the result into two parts (expiration date and email address)
    std::string expirationDate;
    std::string ClassifiedEmail;

    // Optional: Trim extra spaces from the result string
    result.erase(result.find_last_not_of(" \t\r\n") + 1); 
    
    std::stringstream ss(result);
    // Try to read both parts, separated by a comma
    std::getline(ss, expirationDate, ',');  // Assuming result is comma-separated
    std::getline(ss, ClassifiedEmail); // Get the second part as the email

    // Check if both values were found
    if (expirationDate.empty() || ClassifiedEmail.empty()) {
        throw std::runtime_error("PowerShell script did not return expected output (expiration date or email missing).");
    }

    return { expirationDate, ClassifiedEmail };
}

bool IsDuplicateRecordBySerialNumber(const std::wstring& AgencyCardSerialNumber) {
    std::wifstream file(csvPath);
    std::wstring line;

    // Skip the header row (if any)
    std::getline(file, line);

    // Check each line in the CSV file
    while (std::getline(file, line)) {
        // Assuming CSV format is: SerialNumber,Date
        size_t delimiterPos = line.find(L',');
        if (delimiterPos != std::wstring::npos) {
            std::wstring serialNumber = line.substr(0, delimiterPos); // Get the serial number
            if (serialNumber == AgencyCardSerialNumber) {
                return true; // Found a duplicate serial number
            }
        }
    }

    return false; // No duplicate found
}

void EnsureCSVFile(const std::wstring& filePath) {
    // Check if the file exists
    FILE* csvFile = NULL;
    errno_t err = _wfopen_s(&csvFile, filePath.c_str(), L"r, ccs=UTF-8");

    if (err != 0 || csvFile == NULL) {
        // File doesn't exist or failed to open, create the file and write the header

        // Open file in write mode to create it
        err = _wfopen_s(&csvFile, filePath.c_str(), L"w, ccs=UTF-8");

        if (err != 0 || csvFile == NULL) {
            wprintf(L"\n[ERROR] Failed to open CSV file for writing.\n");
            return;  // Exit on error
        }

        // Write the header with two columns: SerialNumber and ExpirationDate
        fwprintf(csvFile, L"SerialNumber,ExpirationDate,Name,ClassifiedEmail\n");
        wprintf(L"\n[INFO] CSV file created and header written: SerialNumber,ExpirationDate,Name,ClassifiedEmail\n");

        // Close the file after writing the header
        fclose(csvFile);
    }
    else {
        // File already exists, no need to write the header again
        fclose(csvFile);
        wprintf(L"\n[INFO] CSV file already exists, no header needed.\n");
    }
}

void WriteToCSVIfNotDuplicate(const std::wstring& filePath, const std::wstring& AgencyCardSerialNumber, const std::string& expirationDate = "", const std::wstring& Name = L"", const std::string& userEmail = "") {
    // Check if the serial number already exists in the file (duplicate check)
    if (!IsDuplicateRecordBySerialNumber(AgencyCardSerialNumber)) {
        FILE* csvFile = NULL;
        errno_t err;
        err = _wfopen_s(&csvFile, filePath.c_str(), L"a, ccs=UTF-8");

        if (err != 0 || csvFile == NULL) {
            wprintf(L"\n[ERROR] Failed to open CSV file for appending.\n");
            return;  // Exit on error
        }

        // Write the new record to the CSV file if it's not a duplicate
        fwprintf(csvFile, L"%ls,%S,%ls,%S\n", AgencyCardSerialNumber.c_str(), expirationDate.c_str(), Name.c_str(), userEmail.c_str());
        wprintf(L"\n[INFO] Data written to CSV: %ls, %S, %ls, %S\n", AgencyCardSerialNumber.c_str(), expirationDate.c_str(), Name.c_str(), userEmail.c_str());

        // Close the file after appending
        fclose(csvFile);
    }
    else {
        wprintf(L"\n[INFO] Duplicate record found, not writing to CSV.\n");
    }
}

// Function to trigger the PowerShell script
void RunPowerShellScript(const std::wstring& scriptPath) {
    // Wait for file release
    WaitForFileRelease(csvPath);

    // Properly quote the script path
    std::wstring command = L"pwsh.exe -ExecutionPolicy Bypass -File \"" + scriptPath + L"\"";

    // Run PowerShell script
    WIN32_FIND_DATA findFileData;
    HANDLE hFind = FindFirstFile(scriptPath.c_str(), &findFileData);
    if (hFind != INVALID_HANDLE_VALUE) {
        // Command to execute the PowerShell script
        system(std::string(command.begin(), command.end()).c_str());
        printf("\n[INFO] PowerShell script executed: %ws\n", scriptPath.c_str());
    }
    else {
        printf("\n[ERROR] PowerShell script not found: %ws\n", scriptPath.c_str());
    }
}

// Function to show a message box and exit the program gracefully
void ShowErrorMessage(const std::wstring& message) {
    MessageBoxW(NULL, message.c_str(), L"Card Reader Error", MB_ICONERROR | MB_OK);
    exit(1);  // Exit after showing the error message
}

// Function to convert month name to month number
int monthToNumber(const std::wstring& month) {
    static std::map<std::wstring, int> monthMap = {
        {L"JAN", 1}, {L"FEB", 2}, {L"MAR", 3}, {L"APR", 4},
        {L"MAY", 5}, {L"JUN", 6}, {L"JUL", 7}, {L"AUG", 8},
        {L"SEP", 9}, {L"OCT", 10}, {L"NOV", 11}, {L"DEC", 12}
    };
    auto it = monthMap.find(month);
    return it != monthMap.end() ? it->second : -1;  // Return -1 if month not found
}

// Function to prompt the user for a new PIN
std::string PromptForPIN() {
    std::string pin;
    std::cout << "Enter a valid PIN: ";
    std::cin >> pin;

    // Ensure PIN is of correct length (for example, 6 digits)
    if (pin.length() != 6) {
        std::cerr << "Invalid PIN length! Please enter a 6-digit PIN." << std::endl;
        return "";
    }

    return pin;
}

// The main function, with the updated logic for scanning multiple cards
int main(int argc, char* argv[]) {
    bool sessionActive = true;
    // Set the locale to the user's default
    _wsetlocale(LC_ALL, L"");
    // Initialize the global variables
    exeDir = GetExecutableDirectory();
    csvPath = exeDir + L"\\CardReaderData.csv";
    scriptPath = exeDir + L"\\SmartCardInventoryUploader.ps1";
    scriptPath2 = exeDir + L"\\certutilExpirationDateSmartCard.ps1";

    LONG lResult;
    SCARDCONTEXT hContext;
    SCARDHANDLE hCard;
    DWORD dwActiveProtocol;
    DWORD dwLen;
    SCARD_IO_REQUEST pIoReq;
    BYTE pBuffer[6000]; // Increased buffer size for larger responses
    USHORT SW;
    DWORD i;
    // Define the default PIN
    std::string defaultPIN = "112233";  // Default PIN for newly provisionned cards via CMS
    std::string enteredPIN;
    // Initialize counter for processed cards
    //int scannedCardCount = 0;

    // APDU to get product name
    BYTE PRODUCT_NAME[] = { 0xFF, 0x70, 0x07, 0x6B, 0x08, 0xA2, 0x06, 0xA0, 0x04,
                            0xA0, 0x02, 0x82, 0x00, 0x00 };

    // APDU to get hardware version
    BYTE HARDWARE_VERSION[] = { 0xFF, 0x70, 0x07, 0x6B, 0x08, 0xA2, 0x06, 0xA0, 0x04,
                                0xA0, 0x02, 0x89, 0x00, 0x00 };

    // APDU to select PIV applet
    BYTE SELECT_PIV_APPLET[] = { 0x00, 0xA4, 0x04, 0x00, 0x09, 0xA0, 0x00, 0x00, 0x03, 0x08, 0x00, 0x00, 0x10, 0x00 };


    // APDU to get printed information
    BYTE GET_PRINTED_INFO[] = { 0x00, 0xCB, 0x3F, 0xFF, 0x05, 0x5C, 0x03, 0x5F, 0xC1, 0x09, 0x00 };

    // Establish context
    lResult = SCardEstablishContext(SCARD_SCOPE_USER, NULL, NULL, &hContext);
    if (SCARD_S_SUCCESS != lResult) {
        printf("\n[ERROR] SCardEstablishContext failed. Error code 0x%08X.\n", lResult);
        return 1;
    }

    // Retrieve the list of available readers
    WCHAR* pmszReaders = NULL;
    DWORD cch = SCARD_AUTOALLOCATE;

    lResult = SCardListReaders(hContext, NULL, (LPWSTR)&pmszReaders, &cch);
    if (SCARD_S_SUCCESS != lResult || pmszReaders == NULL) {
        // No readers found, show error message to the user
        ShowErrorMessage(L"\n[ERROR] No HID OMNIKEY 3021 Card Reader Detected. Please Insert a Reader into the USB Port of Host and relaunch program.");
        return 1;  // Exit after showing the message
    }

    // Initialize SCARD_READERSTATEW for a single reader
    SCARD_READERSTATEW readerState;
    readerState.szReader = pmszReaders;  // Use the first (and only) reader
    readerState.dwEventState = SCARD_STATE_UNAWARE;
    readerState.dwCurrentState = SCARD_STATE_UNAWARE;

    // Monitor for card status change
    printf("\n[INFO] Waiting for card insertion/removal...\n");

    // Store the ATR of the last scanned card
    BYTE previousATR[32];  // To store the previous ATR
    DWORD previousATRLength = 0;
    bool cardAlreadyProcessed = false;  // Flag to track if the current card has already been processed

    while (sessionActive) {
        lResult = SCardGetStatusChangeW(hContext, INFINITE, &readerState, 1);
        if (SCARD_S_SUCCESS != lResult) {
            printf("\n[ERROR] SCardGetStatusChangeW failed. Error code 0x%08X.\n", lResult);
            break;
        }

        // Detect a state change
        if (readerState.dwEventState & SCARD_STATE_CHANGED) {
            // Card inserted
            if ((readerState.dwEventState & SCARD_STATE_PRESENT) && !(readerState.dwCurrentState & SCARD_STATE_PRESENT)) {
                // Card inserted
                printf("\n[INFO] Card inserted!\n");

                // Introduce a small delay to ensure the card is fully initialized and ready
                printf("\n[INFO] Please wait, initializing card...\n");
                //Sleep(6000);  // Delay for 6 seconds

                // Connect to the card
                lResult = SCardConnect(hContext, readerState.szReader, SCARD_SHARE_SHARED, SCARD_PROTOCOL_T0 | SCARD_PROTOCOL_T1, &hCard, &dwActiveProtocol);
                if (SCARD_S_SUCCESS != lResult) {
                    SCardFreeMemory(hContext, pmszReaders);
                    SCardReleaseContext(hContext);
                    printf("\n[ERROR] Cannot detect card. Error code 0x%08X.\n", lResult);
                    return 1;
                }

                // Select protocol based on the card's active protocol
                if (SCARD_PROTOCOL_T1 == dwActiveProtocol) {
                    pIoReq = *SCARD_PCI_T1;
                    printf("\n[INFO] Using protocol T=1\n");
                }
                else if (SCARD_PROTOCOL_T0 == dwActiveProtocol) {
                    pIoReq = *SCARD_PCI_T0;
                    printf("\n[INFO] Using protocol T=0\n");
                }
                else {
                    // In case the protocol is neither T=0 nor T=1
                    printf("\n[ERROR] Unsupported protocol. Error code 0x%08X.\n", lResult);
                    SCardDisconnect(hCard, SCARD_LEAVE_CARD);
                    SCardFreeMemory(hContext, pmszReaders);
                    SCardReleaseContext(hContext);
                    return 1;
                }

                if (SCARD_S_SUCCESS == lResult) {
                    // Get the ATR of the inserted card
                    BYTE currentATR[33];
                    DWORD currentATRLength = sizeof(currentATR);
                    lResult = SCardStatus(hCard, NULL, NULL, NULL, NULL, currentATR, &currentATRLength);
                    if (SCARD_S_SUCCESS != lResult) {
                        printf("\n[ERROR] Failed to get ATR. Error code 0x%08X.\n", lResult);
                        SCardDisconnect(hCard, SCARD_LEAVE_CARD);
                        SCardFreeMemory(hContext, pmszReaders);
                        SCardReleaseContext(hContext);
                        return 1;
                    }

                    // Compare the ATRs
                    if (previousATRLength == currentATRLength && memcmp(previousATR, currentATR, currentATRLength) == 0) {
                        // Card is the same as the previously scanned one, process further
                        printf("\n[ATTENTION] Card already scanned.\n");
                        // Prompt technician to remove the card and insert a new one
                        printf("\n[PROMPT] Please remove the current card and insert a new one to continue scanning.\n");
                        printf("[INFO] Press 'Enter' to continue once the new card is inserted...");
                        // Clear any leftover characters in the input buffer
                        while (getchar() != '\n');  // Consume any leftover characters

                        // Wait for technician input before continuing
                        getchar(); // This will wait for Enter key press

                        // Disconnect the current card to reset the reader
                        SCardDisconnect(hCard, SCARD_LEAVE_CARD);

                        // Store the current ATR as the previous ATR 
                        previousATRLength = currentATRLength;
                        memcpy(previousATR, currentATR, currentATRLength);

                        continue; // Skip to the next iteration of the loop to wait for a new card
                    }

                    // Handle card processing
                    dwLen = sizeof(pBuffer);
                    bool authenticated = false;
                    bool isNewCard = false;
                    std::wstring serialNumber;  // This will hold the serial number as a wide string

                    if (RetrieveSerialNumber(hCard, pIoReq, pBuffer, dwLen, serialNumber)) {
                        // Successfully retrieved the serial number, indicating a brand new card
                        ExtractSerialNumber(pBuffer, dwLen);

                        isNewCard = true;  // Set the flag indicating it's a new card

                        // Handle the new card case: Save serial number to CSV and prompt for next card
                        EnsureCSVFile(csvPath);
                        WriteToCSVIfNotDuplicate(csvPath, serialNumber); // Use L"" for wide string literals
                        scannedCardCount++;

                        // Ask technician to scan another card or exit
                        printf("\n[SUCCESS] Card #%d scanned. Scan another card? (Y/N): ", scannedCardCount);
                        char response;
                        int result = scanf_s(" %c", &response);
                        if (result != 1 || (response != 'Y' && response != 'y')) {
                            sessionActive = false;  // Exit loop if technician doesn't want to continue
                            break;
                        }
                        continue;
                    }
                    else {
                        // Failed to retrieve the serial number, indicating a provisioned card
                        if (VerifyDefaultPIN(hCard, pIoReq, pBuffer, dwLen)) {
                            authenticated = true;
                            enteredPIN = "112233";  // Default PIN
                        }
                        else {
                            enteredPIN = PromptForCustomPIN(hCard, pIoReq, pBuffer, dwLen);
                            authenticated = !enteredPIN.empty();
                        }
                    }
                    if (authenticated && !isNewCard) {
                        try {
                            DWORD dwState = 0;
                            DWORD dwProtocol = 0;
                            BYTE pbAtr[256];
                            DWORD dwAtrLen = sizeof(pbAtr);

                            // Wait until the card is ready to receive APDU commands (SCARD_SPECIFIC state)
                            printf("\n[INFO] Waiting for card to be ready...\n");
                            while (true) {
                                // Get current card status
                                LONG lResult = SCardStatusW(hCard, NULL, NULL, &dwState, &dwProtocol, pbAtr, &dwAtrLen);
                                if (lResult != SCARD_S_SUCCESS) {
                                    printf("[ERROR] Failed to get card status. Error code 0x%08X.\n", lResult);
                                    SCardDisconnect(hCard, SCARD_LEAVE_CARD);
                                    SCardFreeMemory(hContext, pmszReaders);
                                    SCardReleaseContext(hContext);
                                    return 1;
                                }

                                // Check if the card is in SCARD_SPECIFIC state (ready for communication)
                                if (dwState & SCARD_SPECIFIC) {
                                    printf("\n[INFO] Card is ready to receive APDU commands.\n");
                                    break;  // Card is ready, exit the loop
                                }

                                // If not ready, wait for a moment and try again
                                printf("[INFO] Waiting for card to be ready...\n");
                                Sleep(3000);  // Wait for 1 second before checking again
                            }

                            // Now that the card is ready, retrieve printed information
                            printf("\n[INFO] Retrieving Printed Information...\n");
                            // Reset pBuffer and dwLen before transmitting the APDU
                            memset(pBuffer, 0, sizeof(pBuffer));  // Clear pBuffer
                            dwLen = sizeof(pBuffer);  // Set dwLen to the size of pBuffer

                            lResult = SCardTransmit(hCard, &pIoReq, GET_PRINTED_INFO, sizeof(GET_PRINTED_INFO), NULL, pBuffer, &dwLen);
                            if (SCARD_S_SUCCESS != lResult) {
                                printf("\n[ERROR] Failed to get printed information. Error code 0x%08X.\n", lResult);
                                SCardDisconnect(hCard, SCARD_LEAVE_CARD);
                                SCardFreeMemory(hContext, pmszReaders);
                                SCardReleaseContext(hContext);
                                return 1;
                            }

                            SW = pBuffer[dwLen - 2] << 8 | pBuffer[dwLen - 1];
                            if (SW != 0x9000) {
                                printf("\n[ERROR] Failed to retrieve printed information. Error code 0x%04X.\n", SW);
                                SCardDisconnect(hCard, SCARD_LEAVE_CARD);
                                SCardFreeMemory(hContext, pmszReaders);
                                SCardReleaseContext(hContext);
                                return 1;
                            }

                            // Extract Name (Tag 0x01) - Max 125 bytes, Text (ASCII)
                            BYTE tagName = 0x01;
                            wchar_t Name[126] = { 0 };  // Buffer for Name (+1 for null terminator)

                            // Parse and extract name tag
                            for (i = 0; i < dwLen - 2; i++) {
                                if (pBuffer[i] == tagName) { // Found the tag for Agency Card Serial Number
                                    if (i + 1 < dwLen - 2 && pBuffer[i + 1] <= 125) {
                                        MultiByteToWideChar(CP_ACP, 0, (char*)&pBuffer[i + 2], pBuffer[i + 1], Name, 125);
                                        Name[pBuffer[i + 1]] = L'\0'; // Null terminate
                                        break;
                                    }
                                }
                            }

                            // Print the Name
                            wprintf(L"\n[SUCCESS] Name: %ls\n", Name);

                            // Extract Agency Card Serial Number (Tag 0x05)
                            BYTE tagSerial = 0x05;
                            wchar_t AgencyCardSerialNumber[21] = { 0 }; // Buffer for Agency Card Serial Number (+1 for null terminator)

                            for (i = 0; i < dwLen - 2; i++) {
                                if (pBuffer[i] == tagSerial) { // Found the tag for Agency Card Serial Number
                                    if (i + 1 < dwLen - 2 && pBuffer[i + 1] <= 20) {
                                        MultiByteToWideChar(CP_ACP, 0, (char*)&pBuffer[i + 2], pBuffer[i + 1], AgencyCardSerialNumber, 20);
                                        AgencyCardSerialNumber[pBuffer[i + 1]] = L'\0'; // Null terminate
                                        break;
                                    }
                                }
                            }

                            // Print the Agency Card Serial Number
                            wprintf(L"\n[SUCCESS] Agency Card Serial Number: %ls\n", AgencyCardSerialNumber);

                            std::string expirationDate;
                            std::string userEmail;
                            
                            try {
                                // Call RunPowerShellCommand to get both the expiration date and email address
                                std::pair<std::string, std::string> result = RunPowerShellCommand(scriptPath2, enteredPIN);
                                expirationDate = result.first;  // The expiration date is the first element
                                userEmail = result.second;      // The email address is the second element

                                // Output the result from PowerShell (expiration date and email address)
                                if (!expirationDate.empty()) {
                                    std::cout << "\n[SUCCESS] Certificate Expiration Date: " << expirationDate << std::endl;
                                }
                                else {
                                    std::cerr << "\n[ERROR] Failed to extract expiration date from PowerShell script output." << std::endl;
                                }

                                if (!userEmail.empty()) {
                                    std::cout << "\n[SUCCESS] Classified User Email: " << userEmail << std::endl;
                                }
                                else {
                                    std::cerr << "\n[ERROR] Failed to extract classified user email from PowerShell script output." << std::endl;
                                }

                            }
                            catch (const std::exception& e) {
                                std::cerr << "Error running PowerShell command: " << e.what() << std::endl;
                                return 1;
                            }

                            // Now check if the card is a duplicate in the CSV file
                            bool isDuplicate = IsDuplicateRecordBySerialNumber(AgencyCardSerialNumber);
                            if (isDuplicate) {
                                printf("\n[INFO] Card already scanned. Please remove the current card and insert a new one to continue.\n");
                                // Ask if technician wants to continue with the same card
                                printf("\n[PROMPT] Scan another card or continue? (Y/N): ");
                                char response;
                                int result = scanf_s(" %c", &response); // Removed sizeof(response)
                                if (result != 1 || (response != 'Y' && response != 'y')) {
                                    break; // Exit the loop if technician doesn't want to continue
                                }
                                // If 'Y' entered, continue with same card without further processing
                                cardAlreadyProcessed = true;  // Mark that we don't need to process the card info again
                                continue;  // Skip the extraction and CSV saving process for this card
                            }

                            // If no duplicate, write to CSV and process
                            EnsureCSVFile(csvPath);
                            WriteToCSVIfNotDuplicate(csvPath, AgencyCardSerialNumber, expirationDate, Name, userEmail);
           
                            // Increment the scanned card counter
                            scannedCardCount++;

                            // Ask technician to scan another card or exit
                            printf("\n[SUCCESS] Card #%d scanned. Scan another card? (Y/N): ", scannedCardCount);
                            char response;
                            int result = scanf_s(" %c", &response); // Removed sizeof(response)
                            if (result != 1 || (response != 'Y' && response != 'y')) {
                                sessionActive = false;  // Exit loop if technician doesn't want to continue
                                break;
                            }
                        }
                        catch (const std::exception& e) {
                            std::cerr << "Error running PowerShell command: " << e.what() << std::endl;
                            return 1;
                        }
                    }
                }
            }
            // Card removed
            else if ((readerState.dwEventState & SCARD_STATE_EMPTY) && (readerState.dwCurrentState & SCARD_STATE_PRESENT)) {
                // Reset the ATR info when the card is removed
                memset(previousATR, 0, sizeof(previousATR));
                previousATRLength = 0;
                cardAlreadyProcessed = false; // Reset the processing flag
                printf("\n[ATTENTION] Card removed. Please insert a new card.\n");
                break;  // Exit the loop if the card is removed
            }
        }
    }
    // At the end of the session, display the summary
    displayScanSummary();
    // Run PowerShell script to upload the data (if any cards were scanned)
    if (scannedCardCount > 0) {
        printf("\n[INFO] Running PowerShell script to upload data...\n");
        RunPowerShellScript(scriptPath);
    }
    else {
        printf("\n[INFO] No new cards scanned. Exiting...\n");
    }
    // Cleanup and release resources
    SCardFreeMemory(hContext, pmszReaders);
    SCardReleaseContext(hContext);
    return 0;
}