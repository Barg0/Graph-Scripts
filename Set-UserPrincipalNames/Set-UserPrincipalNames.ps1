# Script version:   2025-05-29 15:30
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

# Capture start time to log script duration
$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

# Import CSV Path
$csvPath = Join-Path -Path $PSScriptRoot -ChildPath "UserPrincipalNames.csv"

# ---------------------------[ Script name ]---------------------------

# Script name used for folder/log naming
$scriptName = "Set-UserPrincipalNames"
$logFileName = "$($scriptName).log"

# ---------------------------[ Logging Setup ]---------------------------

# Logging control switches
$log = $true                     # Set to $false to disable logging in shell
$enableLogFile = $false          # Set to $false to disable file output

# Define the log output location
$logFileDirectory = "$PSScriptRoot"
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

# ---------------------------[ Graph Connection ]---------------------------

# Ensure the required Graph module is available
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Write-Log "Microsoft.Graph.Users module not found. Installing..." -Tag "Info"
    try {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -ErrorAction Stop
        Write-Log "Microsoft.Graph module installed successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to install Microsoft.Graph: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
} else {
    Write-Log "Microsoft.Graph.Users module found." -Tag "Info"
}

# Import the required module
try {
    Import-Module Microsoft.Graph.Users -Force -ErrorAction Stop
    Write-Log "Microsoft.Graph.Users module imported." -Tag "Success"
} catch {
    Write-Log "Failed to import Microsoft.Graph.Users module: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

$connected = $false
try {
    $context = Get-MgContext
    if ($null -ne $context.Account -and $null -ne $context.Scopes -and $context.Scopes -contains "User.ReadWrite.All") {
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
        Connect-MgGraph -Scopes "User.ReadWrite.All" | Out-Null
        Write-Log "Connected to Microsoft Graph successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

# ---------------------------[ Read CSV ]---------------------------


if (-not (Test-Path $csvPath)) {
    Write-Log "CSV file not found: $csvPath" -Tag "Error"
    Complete-Script -ExitCode 1
}

try {
    $entries = Import-Csv -Path $csvPath
    Write-Log "Loaded $($entries.Count) entries from CSV." -Tag "Info"
} catch {
    Write-Log "Failed to read CSV: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

# ---------------------------[ Update UPNs ]---------------------------

foreach ($entry in $entries) {
    $oldUPN = $entry.Old.Trim()
    $newUPN = $entry.New.Trim()

    if ([string]::IsNullOrWhiteSpace($oldUPN) -or [string]::IsNullOrWhiteSpace($newUPN)) {
        Write-Log "Skipping entry with missing values: '$oldUPN' -> '$newUPN'" -Tag "Error"
        continue
    }

    Write-Log "Attempting UPN change: '$oldUPN' -> '$newUPN'" -Tag "Info"

    try {
        $user = Get-MgUser -UserId $oldUPN -ErrorAction Stop

        if ($user.UserPrincipalName -ieq $newUPN) {
            Write-Log "UPN already set to '$newUPN'. Skipping." -Tag "Info"
            continue
        }

        Update-MgUser -UserId $user.Id -UserPrincipalName $newUPN -ErrorAction Stop
        Write-Log "Successfully updated UPN to '$newUPN'." -Tag "Success"
    } catch {
        Write-Log "Failed to update '$oldUPN' to '$newUPN'. Error: $_" -Tag "Error"
    }
}

# ---------------------------[ Completion ]---------------------------

Complete-Script -ExitCode 0
