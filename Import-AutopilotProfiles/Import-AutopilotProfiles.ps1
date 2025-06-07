# Script version:   2025-06-07 18:33
# Script author:    Barg0

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$graphScopes = @(
    "DeviceManagementServiceConfig.ReadWrite.All"
)
$uri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles"

$scriptName = "Import-AutopilotProfiles"
$logFileName = "$($scriptName).log"

# Output folder for JSON
$importDir = Join-Path -Path $PSScriptRoot -ChildPath "Import"

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

# ---------------------------[ Load and Import JSON Files ]---------------------------

if (-not (Test-Path -Path $importDir)) {
    Write-Log "Import folder not found: $importDir" -Tag "Error"
    Complete-Script -ExitCode 1
}

$jsonFiles = Get-ChildItem -Path $importDir -Filter *.json
if ($jsonFiles.Count -eq 0) {
    Write-Log "No JSON files found in $importDir" -Tag "Info"
    Complete-Script -ExitCode 0
}

Write-Log "Graph URI: $uri" -Tag "Info"

# Get existing profile names to avoid duplicates
Write-Log "Querying existing Autopilot deployment profiles..." -Tag "Info"
try {
    $existingProfilesResponse = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    $existingProfileNames = $existingProfilesResponse.value.displayName
} catch {
    Write-Log "Failed to retrieve existing profiles: $($_.Exception.Message)" -Tag "Error"
    Complete-Script -ExitCode 1
}

foreach ($file in $jsonFiles) {
    # Write-Log "Processing: $($file.Name)" -Tag "Debug"

    try {
        $profileObject = Get-Content -Raw -Path $file.FullName | ConvertFrom-Json
    } catch {
        Write-Log "Failed to parse JSON: $($_.Exception.Message)" -Tag "Error"
        continue
    }

    $displayName = $profileObject.displayName.Trim()

    if ($existingProfileNames -contains $displayName) {
        Write-Log "Profile '$displayName' already exists. Skipping import." -Tag "Info"
        continue
    }

    # Remove read-only and identity properties
    $profileObject.PSObject.Properties.Remove("id")
    $profileObject.PSObject.Properties.Remove("createdDateTime")
    $profileObject.PSObject.Properties.Remove("lastModifiedDateTime")
    $profileObject.PSObject.Properties.Remove("managementServiceAppId")

    try {
        $body = $profileObject | ConvertTo-Json -Depth 20
        $headers = @{ "Content-Type" = "application/json" }

        Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -Headers $headers -ErrorAction Stop | Out-Null
        Write-Log "Successfully imported: $($file.Name)" -Tag "Success"
    } catch {
        Write-Log "Failed to import $($file.Name): $($_.Exception.Message)" -Tag "Error"
    }
}

# ---------------------------[ Script Completion ]---------------------------

Complete-Script -ExitCode 0