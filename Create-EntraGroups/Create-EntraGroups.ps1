# ====================================================================
# Script version:   2025-05-29 11:10
# Script author:    Barg0
# ====================================================================

# ---------------------------[ Script Start Timestamp ]---------------------------
$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$CsvPath = ".\Groups.csv"

# ---------------------------[ Script Configuration ]---------------------------

$scriptName = "CreateEntraGroups"
$logFileName = "$($logFileName).log"

# ---------------------------[ Logging Setup ]---------------------------

# Logging control switches
$log = $true                     # Set to $false to disable logging
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

# ---------------------------[ Import Modules ]---------------------------
try {
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop
    Write-Log "Microsoft.Graph.Groups module imported successfully." -Tag "Success"
} catch {
    Write-Log "Failed to import Microsoft.Graph.Groups module: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

# ---------------------------[ Connect to Microsoft Graph ]---------------------------
try {
    Write-Log "Connecting to Microsoft Graph..." -Tag "Info"
    Connect-MgGraph -Scopes "Group.ReadWrite.All" | Out-Null
    Write-Log "Connected to Microsoft Graph successfully." -Tag "Success"
} catch {
    Write-Log "Authentication with Microsoft Graph failed: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

# ---------------------------[ Load and Validate CSV ]---------------------------
if (-not (Test-Path $CsvPath)) {
    Write-Log "CSV file not found at path: $CsvPath" -Tag "Error"
    Complete-Script -ExitCode 1
}

try {
    $csvData = Import-Csv -Path $CsvPath
    Write-Log "CSV file loaded successfully." -Tag "Success"
} catch {
    Write-Log "Failed to read CSV: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

# ---------------------------[ Process Each Group ]---------------------------
foreach ($entry in $csvData) {
    $groupType = $entry.GroupType.ToUpper()
    $groupName = $entry.GroupName
    $dynamicRule = $entry.DynamicRule

    if (-not $groupType -or -not $groupName) {
        Write-Log "Missing GroupType or GroupName in a row. Skipping." -Tag "Error"
        continue
    }

    $mailNickname = ($groupName -replace '[^a-zA-Z0-9]', '')

    Write-Log "Creating group '$groupName' of type '$groupType'" -Tag "Info"

    try {
        switch ($groupType) {
            "DU" {
                if (-not $dynamicRule) {
                    Write-Log "DynamicRule is missing for group type 'DU'. Skipping." -Tag "Error"
                    continue
                }

                $null = New-MgGroup -DisplayName $groupName `
                                    -Description "Dynamic User Group: $groupName" `
                                    -GroupTypes @("DynamicMembership") `
                                    -MailEnabled:$false `
                                    -MailNickname $mailNickname `
                                    -SecurityEnabled:$true `
                                    -MembershipRule $dynamicRule `
                                    -MembershipRuleProcessingState "On"
            }

            "DD" {
                if (-not $dynamicRule) {
                    Write-Log "DynamicRule is missing for group type 'DD'. Skipping." -Tag "Error"
                    continue
                }

                $null = New-MgGroup -DisplayName $groupName `
                                    -Description "Dynamic Device Group: $groupName" `
                                    -GroupTypes @("DynamicMembership") `
                                    -MailEnabled:$false `
                                    -MailNickname $mailNickname `
                                    -SecurityEnabled:$true `
                                    -MembershipRule $dynamicRule `
                                    -MembershipRuleProcessingState "On"
            }

            "A" {
                $null = New-MgGroup -DisplayName $groupName `
                                    -Description "Assigned Security Group: $groupName" `
                                    -GroupTypes @() `
                                    -MailEnabled:$false `
                                    -MailNickname $mailNickname `
                                    -SecurityEnabled:$true
            }

            Default {
                Write-Log "Invalid GroupType '$groupType' for group '$groupName'. Skipping." -Tag "Error"
                continue
            }
        }

        Write-Log "Group '$groupName' of type '$groupType' created successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to create group '$groupName': $_" -Tag "Error"
    }
}

# ---------------------------[ Script Completion ]---------------------------
Complete-Script -ExitCode 0
