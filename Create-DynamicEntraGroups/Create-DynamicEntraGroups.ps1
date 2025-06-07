# Script version:   2025-06-07 12:30
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$graphScopes = @(
    "Group.ReadWrite.All"
    )
$uri = "https://graph.microsoft.com/v1.0/groups"

$scriptName = "Create-DynamicEntraGroups"
$logFileName = "$scriptName.log"

# Import csv path
$csvPath = Join-Path -Path $PSScriptRoot -ChildPath "Groups.csv"

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
} else {
    Write-Log "Microsoft.Graph module is already loaded." -Tag "Info"
}

# ---------------------------[ Graph Authentication ]---------------------------

$connected = $false
try {
    $context = Get-MgContext
    if ($null -ne $context.Account -and $null -ne $context.Scopes -and $graphScopes | ForEach-Object { $_ -in $context.Scopes }) {
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

# ---------------------------[ Load CSV ]---------------------------

if (-not (Test-Path $csvPath)) {
    Write-Log "CSV file not found at path: $csvPath" -Tag "Error"
    Complete-Script -ExitCode 1
}

try {
    $csvData = Import-Csv -Path $csvPath
    Write-Log "CSV file loaded with $($csvData.Count) rows." -Tag "Success"
} catch {
    Write-Log "Failed to load CSV: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

# ---------------------------[ Create Groups ]---------------------------

foreach ($entry in $csvData) {
    $groupName   = $entry.GroupName
    $dynamicRule = $entry.DynamicRule
    $description = $entry.Description

    if (-not $groupName -or -not $dynamicRule) {
        Write-Log "Missing GroupName or DynamicRule in row. Skipping." -Tag "Error"
        continue
    }

    $mailNickname = ($groupName -replace '[^a-zA-Z0-9]', '')

    # Check if the group already exists by displayName
    $escapedGroupName = $groupName.Replace("'", "''")
    $checkUri = "$($uri)?`$filter=displayName eq '$escapedGroupName'"

    try {
        $existingGroupResponse = Invoke-MgGraphRequest -Method GET -Uri $checkUri
        if ($existingGroupResponse.value.Count -gt 0) {
            Write-Log "Group '$groupName' already exists. Skipping creation." -Tag "Info"
            continue
        }
    } catch {
        Write-Log "Failed to check existence of group '$groupName': $_" -Tag "Error"
        continue
    }

    $payload = @{
        displayName                    = $groupName
        description                    = $description
        mailEnabled                    = $false
        mailNickname                   = $mailNickname
        securityEnabled                = $true
        groupTypes                     = @("DynamicMembership")
        membershipRule                 = $dynamicRule
        membershipRuleProcessingState = "On"
    } | ConvertTo-Json -Depth 3

    try {
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body $payload -ContentType "application/json" | Out-Null
        Write-Log "Group '$groupName' created successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to create group '$groupName': $_" -Tag "Error"
    }
}


# ---------------------------[ Done ]---------------------------

Complete-Script -ExitCode 0
