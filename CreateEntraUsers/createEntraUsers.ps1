# Script version:   2025-06-01 10:40
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

# Capture start time to log script duration
$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

# Import CSV Path
$importCSV = Join-Path -Path $PSScriptRoot -ChildPath "import.csv"
# Created Users/Password export path
$exportCSV = Join-Path -Path $PSScriptRoot -ChildPath "users.csv"

# Define the length of the Password for the Users
$passwordLength = 14
# Does the user need to change his password on first sign-in? 
$passwordChange = $false

# ---------------------------[ Script name ]---------------------------

# Script name used for folder/log naming
$scriptName = "createEntraUsers"
$logFileName = "$($scriptName).log"

# ---------------------------[ Logging Setup ]---------------------------

# Logging control switches
$log = $true                     # Set to $false to disable logging in shell
$enableLogFile = $false          # Set to $false to disable file output

# Define the log output location
$logFileDirectory = "$env:ProgramData"
$logFile = "$logFileDirectory\$logFileName"

# Ensure the log directory exists
if ($enableLogFile -and -not (Test-Path $logFileDirectory)) {
    New-Item -ItemType Directory -Path $logFileDirectory -Force | Out-Null
}

# Function to write structured logs to file and console
function Write-Log {
    param ([string]$Message, [string]$Tag = "Info")

    if (-not $log) { return } # Exit if logging is disabled

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $tagList = @("Start", "Check", "Info", "Success", "Error", "Debug", "End")
    $rawTag = $Tag.Trim()

    if ($tagList -contains $rawTag) {
        $rawTag = $rawTag.PadRight(7)
    } else {
        $rawTag = "Error  "  # Fallback if an unrecognized tag is used
    }

    # Set tag colors
    $color = switch ($rawTag.Trim()) {
        "Start"   { "Cyan" }
        "Check"   { "Blue" }
        "Info"    { "Yellow" }
        "Success" { "Green" }
        "Error"   { "Red" }
        "Debug"   { "DarkYellow"}
        "End"     { "Cyan" }
        default   { "White" }
    }

    $logMessage = "$timestamp [  $rawTag ] $Message"

    # Write to file if enabled
    if ($enableLogFile) {
        "$logMessage" | Out-File -FilePath $logFile -Append
    }

    # Write to console with color formatting
    Write-Host "$timestamp " -NoNewline
    Write-Host "[  " -NoNewline -ForegroundColor White
    Write-Host "$rawTag" -NoNewline -ForegroundColor $color
    Write-Host " ] " -NoNewline -ForegroundColor White
    Write-Host "$Message"
}

# ---------------------------[ Exit Function ]---------------------------

function Complete-Script {
    param([int]$ExitCode)
    $scriptEndTime = Get-Date
    $duration = $scriptEndTime - $scriptStartTime
    Write-Log "Script execution time: $($duration.ToString("hh\:mm\:ss\.ff"))" -Tag "Info"
    Write-Log "Exit Code: $ExitCode" -Tag "Info"
    Write-Log "======== Script Completed ========" -Tag "End"
    exit $ExitCode
}
# Complete-Script -ExitCode 0

# ---------------------------[ Script Start ]---------------------------

Write-Log "======== Script Started ========" -Tag "Start"
Write-Log "ComputerName: $env:COMPUTERNAME | User: $env:USERNAME | Script: $scriptName" -Tag "Info"

# ---------------------------[ Password Generator ]---------------------------
function New-SecureRandomPassword {
    $allowed = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ23456789!@#$%^&*()-_=+'.ToCharArray()
    do {
        $plainPassword = -join ((1..$passwordLength) | ForEach-Object { $allowed | Get-Random })
    } while (
        ($plainPassword -notmatch '[A-Z]') -or
        ($plainPassword -notmatch '[a-z]') -or
        ($plainPassword -notmatch '\d') -or
        ($plainPassword -notmatch '[!@#$%^&*()\-_=+]')
    )
    # Write-Log "Generated secure random password." -Tag "Success"
    return @{
        SecurePassword = (ConvertTo-SecureString $plainPassword -AsPlainText -Force)
        PlainText = $plainPassword
    }
}

# ---------------------------[ Convert CSV Row to Parameters ]---------------------------
function Convert-CsvToUserParams {
    param ($user)
    $params = @{}

    if ($user.UserPrincipalName) { $params["UserPrincipalName"] = $user.UserPrincipalName }
    if ($user.AccountEnabled) { $params["AccountEnabled"] = [bool]::Parse($user.AccountEnabled) }
    if ($user.DisplayName) { $params["DisplayName"] = $user.DisplayName }
    if ($user.UserPrincipalName) { $params["MailNickname"] = ($user.UserPrincipalName -split "@")[0] }

    if ($user.FirstName) { $params["GivenName"] = $user.FirstName }
    if ($user.LastName)  { $params["Surname"]   = $user.LastName }
    if ($user.JobTitle)  { $params["JobTitle"]  = $user.JobTitle }
    if ($user.Department) { $params["Department"] = $user.Department }
    if ($user.CompanyName) { $params["CompanyName"] = $user.CompanyName }

    if ($user.StreetAddress)  { $params["StreetAddress"] = $user.StreetAddress }
    if ($user.City)           { $params["City"] = $user.City }
    if ($user.StateOrProvince){ $params["State"] = $user.StateOrProvince }
    if ($user.PostalCode)     { $params["PostalCode"] = $user.PostalCode }
    if ($user.Country)        { 
        $params["Country"] = $user.Country
        $params["UsageLocation"] = $user.Country
    } else {
        $params["UsageLocation"] = "DE"
    }

    if ($user.BusinessPhone) { $params["BusinessPhones"] = @($user.BusinessPhone) }
    if ($user.MobilePhone)   { $params["MobilePhone"] = $user.MobilePhone }
    if ($user.FaxNumber)     { $params["FaxNumber"] = $user.FaxNumber }

    return $params
}

# ---------------------------[ Create User Function ]---------------------------
function New-EntraUserFromCsvRow {
    param ($csvRow)

    try {
        $passwordObj = New-SecureRandomPassword

        $userParams = Convert-CsvToUserParams -user $csvRow
        $userParams["PasswordProfile"] = @{
            Password = $passwordObj.PlainText
            ForceChangePasswordNextSignIn = $passwordChange
        }

        New-MgUser @userParams | Out-Null
        Write-Log "Successfully created user: $($csvRow.UserPrincipalName)" -Tag "Success"

        # Append UPN and password to password export CSV
        [PSCustomObject]@{
            UserPrincipalName = $csvRow.UserPrincipalName
            Password          = $passwordObj.PlainText
        } | Export-Csv -Path $exportCSV -Append -NoTypeInformation -Encoding UTF8

    } catch {
        Write-Log "Error creating user $($csvRow.UserPrincipalName): $_" -Tag "Error"
    }
}

# ---------------------------[ Main Execution ]---------------------------
# Ensure Graph SDK is available
if (-not (Get-Module Microsoft.Graph.Users -ListAvailable)) {
    Write-Log "Installing Microsoft.Graph module..." -Tag "Info"
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph.Users

try {
    Write-Log "Connecting to Microsoft Graph..." -Tag "Info"
    Connect-MgGraph -Scopes "User.ReadWrite.All" | Out-Null
} catch {
    Write-Log "Failed to connect to Microsoft Graph: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

# Check if CSV is present
if (-not (Test-Path $importCSV)) {
    Write-Log "CSV file not found at $importCSV" -Tag "Error"
    Complete-Script -ExitCode 1
}

# Initialize output CSV with headers
if (Test-Path $exportCSV) { Remove-Item $exportCSV -Force }
"UserPrincipalName,Password" | Out-File -FilePath $exportCSV -Encoding UTF8

# Process each user
$csvData = Import-Csv -Path $importCSV
foreach ($user in $csvData) {
    New-EntraUserFromCsvRow -csvRow $user
}

Complete-Script -ExitCode 0