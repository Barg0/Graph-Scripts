# Script version:   2025-06-08
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$graphScopes = @(
    "Application.Read.All",
    "Policy.Read.All",
    "Policy.ReadWrite.ConditionalAccess",
    "Directory.Read.All",
    "User.Read.All"
)

$scriptName = "Import-ConditionalAccessPolicies"
$logFileName = "$($scriptName).log"

$scriptStartTime = Get-Date

# ---------------------------[ Logging Control ]---------------------------

$log = $true                     # Set to $false to disable logging in shell
$enableLogFile = $false          # Set to $false to disable file output

# Define the log output location
$logFileDirectory = "$PSScriptRoot"
$logFile = Join-Path -Path $logFileDirectory -ChildPath $logFileName

# ---------------------------[ Logging Setup ]---------------------------

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

# ---------------------------[ Graph SDK Setup ]---------------------------

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Log "Microsoft.Graph module not found. Installing..." -Tag "Info"
    try {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -ErrorAction Stop
        Write-Log "Microsoft.Graph module installed successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to install Microsoft.Graph: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
} else {
    Write-Log "Microsoft.Graph module found." -Tag "Info"
}

if (-not (Get-Module Microsoft.Graph)) {
    try {
        Import-Module Microsoft.Graph -Force -ErrorAction Stop
        Write-Log "Microsoft.Graph module imported." -Tag "Success"
    } catch {
        Write-Log "Failed to import Microsoft.Graph module: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

# ---------------------------[ Graph Authentication ]---------------------------

$connected = $false
try {
    $context = Get-MgContext
    if ($null -ne $context.Account -and $null -ne $context.Scopes -and $context.Scopes -contains $graphScopes) {
        Write-Log "Microsoft Graph already connected as $($context.Account)" -Tag "Success"
        $connected = $true
    } else {
        Write-Log "Microsoft Graph context incomplete or lacks required scope. Reconnecting..." -Tag "Info"
    }
} catch {
    Write-Log "Microsoft Graph not connected. Attempting connection..." -Tag "Info"
}

if (-not $connected) {
    try {
        Connect-MgGraph -Scopes $graphScopes | Out-Null
        Write-Log "Connected to Microsoft Graph successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

# ---------------------------[ Prompt Entra P1/P2 Selection ]---------------------------

$validLicenseSelected = $false
while (-not $validLicenseSelected) {
    $entraLicense = Read-Host "Which Entra License do you have? Enter '1' for Entra P1 (E3/Business Premium) or '2' for Entra P2 (E5)"

    switch ($entraLicense) {
        "1" {
            Write-Log "Selected Entra P1 (E3 / Business Premium)" -Tag "Info"
            $entraLicense = "P1"
            $validLicenseSelected = $true
        }
        "2" {
            Write-Log "Selected Entra P2 (E5)" -Tag "Info"
            $entraLicense = "P2"
            $validLicenseSelected = $true
        }
        default {
            Write-Log "Invalid selection. Please enter '1' for P1 or '2' for P2." -Tag "Error"
        }
    }
}

# ---------------------------[ Prompt for Emergency Access UPN ]---------------------------

$emergencyUpn = Read-Host "Enter the UPN of the emergency access account"
try {
    $emergencyUser = Get-MgUser -UserId $emergencyUpn -ErrorAction Stop
    $emergencyAccessAccountId = $emergencyUser.Id
    Write-Log "Found user '$($emergencyUser.DisplayName)' with ID $emergencyAccessAccountId" -Tag "Success"
} catch {
    Write-Log "Failed to find user with UPN '$emergencyUpn': $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

# ---------------------------[ Helpers for object creation ]---------------------------

function New-Group {
    param([string]$DisplayName)

    $group = Get-MgGroup -Filter "displayName eq '$DisplayName'" -ConsistencyLevel eventual -CountVariable count
    if ($group) {
        Write-Log "Group '$DisplayName' already exists." -Tag "Info"
        return $group.Id
    }

    Write-Log "Creating group '$DisplayName'" -Tag "Info"
    $groupBody = @{
        displayName     = $DisplayName
        mailEnabled     = $false
        mailNickname    = ($DisplayName -replace ' ', '')
        securityEnabled = $true
        groupTypes      = @()
    }

    $newGroup = New-MgGroup -BodyParameter $groupBody
    Write-Log "Group '$DisplayName' created with ID: $($newGroup.Id)" -Tag "Success"
    return $newGroup.Id
}

function GetOrCreateCountryNamedLocation {
    param ([string]$Name, [string[]]$Countries)

    $existing = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identity/conditionalAccess/namedLocations").value |
        Where-Object { $_.displayName -eq $Name -and $_.'@odata.type' -eq '#microsoft.graph.countryNamedLocation' }

    if ($existing) {
        Write-Log "Named location '$Name' already exists with ID: $($existing.id)" -Tag "Info"
        return $existing.id
    }

    $body = @{
        "@odata.type" = "#microsoft.graph.countryNamedLocation"
        displayName = $Name
        countriesAndRegions = $Countries
        includeUnknownCountriesAndRegions = $false
    }

    $location = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/identity/conditionalAccess/namedLocations" -Body ($body | ConvertTo-Json -Depth 3)
    Write-Log "Named location '$Name' created with ID: $($location.id)" -Tag "Success"
    return $location.id
}

function New-CountryNamedLocation {
    param ([string]$Name, [string[]]$Countries)

    $existing = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identity/conditionalAccess/namedLocations").value |
        Where-Object { $_.displayName -eq $Name -and $_.'@odata.type' -eq '#microsoft.graph.countryNamedLocation' }

    if ($existing) {
        Write-Log "Named location '$Name' already exists with ID: $($existing.id)" -Tag "Info"
        return $existing.id
    }

    $body = @{
        "@odata.type" = "#microsoft.graph.countryNamedLocation"
        displayName = $Name
        countriesAndRegions = $Countries
        includeUnknownCountriesAndRegions = $false
    }

    $location = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/identity/conditionalAccess/namedLocations" -Body ($body | ConvertTo-Json -Depth 3)
    Write-Log "Named location '$Name' created with ID: $($location.id)" -Tag "Success"
    return $location.id
}

function New-IpNamedLocation {
    param (
        [string]$Name,
        [string[]]$IpRanges,
        [switch]$IsTrusted
    )

    $existing = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identity/conditionalAccess/namedLocations").value |
        Where-Object { $_.displayName -eq $Name -and $_.'@odata.type' -eq '#microsoft.graph.ipNamedLocation' }

    if ($existing) {
        Write-Log "IP Named location '$Name' already exists with ID: $($existing.id)" -Tag "Info"
        return $existing.id
    }

    # Properly format ipRanges array with required @odata.type
    $formattedRanges = @()
    foreach ($range in $IpRanges) {
        $formattedRanges += @{
            "@odata.type" = "#microsoft.graph.iPv4CidrRange"
            "cidrAddress" = $range
        }
    }

    $body = @{
        "@odata.type" = "#microsoft.graph.ipNamedLocation"
        "displayName" = $Name
        "isTrusted"   = $IsTrusted.IsPresent
        "ipRanges"    = $formattedRanges
    }

    try {
        $response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/identity/conditionalAccess/namedLocations" -Body ($body | ConvertTo-Json -Depth 10)
        Write-Log "IP Named location '$Name' created with ID: $($response.id)" -Tag "Success"
        return $response.id
    } catch {
        Write-Log "Failed to create IP Named location '$Name': $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

function Test-AuthenticationContext {
    $displayName = "Privileged Role Activation"
    $customId = "c1"

    $existing = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identity/conditionalAccess/authenticationContextClassReferences").value |
        Where-Object { $_.displayName -eq $displayName }

    if ($null -ne $existing) {
        Write-Log "Authentication Context '$displayName' already exists with ID: $($existing.id)" -Tag "Info"
        return $existing.id
    }

    Write-Log "Creating Authentication Context '$displayName' with ID '$customId'" -Tag "Info"

    $body = @{
        id = $customId
        displayName = $displayName
        description = "Privileged role elevation protection"
        isAvailable = $true
    }

    try {
        $created = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/identity/conditionalAccess/authenticationContextClassReferences" -Body ($body | ConvertTo-Json -Depth 3)
        Write-Log "Authentication Context '$displayName' created with ID: $($created.id)" -Tag "Success"
        return $created.id
    } catch {
        Write-Log "Failed to create Authentication Context '$displayName': $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

function Test-TemporaryAccessPassAuthStrength {
    $displayName = "Temporary Access Pass"
    $description = "Requires a one-time use Temporary Access Pass for secure and time-limited authentication scenarios."

    $existing = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/policies/authenticationStrengthPolicies").value |
        Where-Object { $_.displayName -eq $displayName }

    if ($null -ne $existing) {
        Write-Log "Authentication Strength '$displayName' already exists with ID: $($existing.id)" -Tag "Info"
        return $existing.id
    }

    Write-Log "Creating Authentication Strength '$displayName'" -Tag "Info"

    $body = @{
        displayName = $displayName
        description = $description
        policyType = "custom"
        allowedCombinations = @("temporaryAccessPassOneTime")
    }

    try {
        $created = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/policies/authenticationStrengthPolicies" -Body ($body | ConvertTo-Json -Depth 3)
        Write-Log "Authentication Strength '$displayName' created with ID: $($created.id)" -Tag "Success"
        return $created.id
    } catch {
        Write-Log "Failed to create Authentication Strength '$displayName': $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}
# ---------------------------[ Object Creation ]---------------------------

$azureExclusionsGroupId = New-Group -DisplayName "Conditional Access - Microsoft Azure exclusions"
$entraServiceAccountsGroupId = New-Group -DisplayName "Conditional Access - Service accounts"
$entraAppProtectionGroupId = New-Group -DisplayName "Conditional Access - App protection"
$travelGroupId = New-Group -DisplayName "Conditional Access - Travel"
$countriesWhitelistId = New-CountryNamedLocation -Name "Countries - Whitelist" -Countries @("DE", "LU")
$countriesTravelId = New-CountryNamedLocation -Name "Countries - Travel" -Countries @("DE")
$countriesGuestsId = New-CountryNamedLocation -Name "Countries - Guests" -Countries @("DE")
New-IpNamedLocation -Name "IP ranges - Company" -IpRanges @("10.10.10.10/32") -IsTrusted
$temporaryAccessPassId = Test-TemporaryAccessPassAuthStrength

if ($entraLicense -eq "P2")
{
    $eligibleAdminsGroupId = New-Group -DisplayName "Conditional Access - Eligible admins"
    Test-AuthenticationContext
} else {
}

if ($entraLicense -eq "P1") {
    $importDir = Join-Path -Path $PSScriptRoot -ChildPath "EntraP1"
    Write-Log "Policy Folder: $importDir" -Tag "Info"
} elseif ($entraLicense -eq "P2") {
    $importDir = Join-Path -Path $PSScriptRoot -ChildPath "EntraP2"
    Write-Log "Policy Folder: $importDir" -Tag "Info"
}

# ---------------------------[ Import Policies ]---------------------------

$existingPolicies = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identity/conditionalAccess/policies").value

Get-ChildItem -Path $importDir -Filter *.json | ForEach-Object {
    $file = $_.FullName
    $json = Get-Content -Path $file -Raw | ConvertFrom-Json

    if ($existingPolicies.displayName -contains $json.displayName) {
        Write-Log "Skipping import: Policy '$($json.displayName)' already exists." -Tag "Info"
        return
    }

    $json = $json | ConvertTo-Json -Depth 10
    $json = $json -replace "<EmergencyAccessAccount>", $emergencyAccessAccountId
    $json = $json -replace "<EligibleAdminsGroup>", $eligibleAdminsGroupId
    $json = $json -replace "<MicrosoftAzureExclusions>", $azureExclusionsGroupId
    $json = $json -replace "<CountriesWhitelist>", $countriesWhitelistId
    $json = $json -replace "<CountriesTravel>", $countriesTravelId
    $json = $json -replace "<CountriesGuests>", $countriesGuestsId
    $json = $json -replace "<TemporaryAccessPass>", $temporaryAccessPassId
    $json = $json -replace "<ServiceAccountsGroup>", $entraServiceAccountsGroupId
    $json = $json -replace "<AppProtectionGroup>", $entraAppProtectionGroupId
    $json = $json -replace "<TravelGroup>", $travelGroupId

    try {
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/identity/conditionalAccess/policies" -Body $json -ErrorAction Stop | Out-Null
        Write-Log "Successfully imported policy '$($file)'." -Tag "Success"
    } catch {
        Write-Log "Failed to import policy '$($file)': $_" -Tag "Error"
    }
}

Complete-Script -ExitCode 0