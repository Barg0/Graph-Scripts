# Script version:   2025-06-04 13:33
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

# Capture start time to log script duration
$scriptStartTime = Get-Date

# ---------------------------[ Paramter ]---------------------------

$graphScopes = @(

    "Policy.ReadWrite.ConditionalAccess",
    "Application.Read.All"
    "User.Read.All"
)
$graphUri = "https://graph.microsoft.com/beta/identity/conditionalAccess/policies"
$importDir = Join-Path -Path $PSScriptRoot -ChildPath "Import"

# ---------------------------[ Script name ]---------------------------

# Script name used for folder/log naming
$scriptName = "Import-ConditionalAccessPolicies"
$logFileName = "$($scriptName)"

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
    if ($null -ne $context.Account -and $null -ne $context.Scopes -and $context.Scopes -contains "Policy.ReadWrite.ConditionalAccess") {
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

# ---------------------------[ Prompt for UserPrincipalName ]---------------------------

$userPrincipalName = Read-Host "Enter the UPN of the user to exclude from all policies"
try {
    $userObject = Get-MgUser -UserId $userPrincipalName -ErrorAction Stop
    $excludedUserId = $userObject.Id
    Write-Log "User found: $($userObject.DisplayName) [$excludedUserId]" -Tag "Success"
} catch {
    Write-Log "Failed to resolve user '$userPrincipalName': $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

# ---------------------------[ Validate Import Folder ]---------------------------

if (-not (Test-Path -Path $importDir)) {
    Write-Log "Import directory not found: $importDir" -Tag "Error"
    Complete-Script -ExitCode 1
}

# ---------------------------[ Import Each Policy ]---------------------------

$files = Get-ChildItem -Path $importDir -Filter *.json
if ($files.Count -eq 0) {
    Write-Log "No JSON files found in $importDir" -Tag "Error"
    Complete-Script -ExitCode 1
}

Write-Log "Found $($files.Count) policy files to import." -Tag "Info"

foreach ($file in $files) {
    Write-Log "Processing file: $($file.Name)" -Tag "Info"

    try {
        $rawContent = Get-Content -Path $file.FullName -Raw -ErrorAction Stop
        $policy = $rawContent | ConvertFrom-Json -ErrorAction Stop
    } catch {
        Write-Log "Failed to read or parse JSON from $($file.Name): $_" -Tag "Error"
        continue
    }

    # ---------------------------[ Modify Policy ]---------------------------

    # Remove read-only Graph-managed properties
    $nullProps = @("id", "createdDateTime", "modifiedDateTime")
    foreach ($prop in $nullProps) {
        if ($policy.PSObject.Properties.Name -contains $prop) {
            $policy.PSObject.Properties.Remove($prop)
        }
    }

    # Enforce report-only mode
    $policy.state = "temporaryAccessPassOneTime"

    # Ensure excludeUsers includes specified user
    if (-not $policy.conditions) {
        $policy | Add-Member -MemberType NoteProperty -Name conditions -Value @{}
    }
    if (-not $policy.conditions.users) {
        $policy.conditions | Add-Member -MemberType NoteProperty -Name users -Value @{}
    }
    if (-not $policy.conditions.users.excludeUsers) {
        $policy.conditions.users.excludeUsers = @()
    }

    if ($policy.conditions.users.excludeUsers -notcontains $excludedUserId) {
        $policy.conditions.users.excludeUsers += $excludedUserId
        Write-Log "Added user to excludeUsers: $excludedUserId" -Tag "Debug"
    }

    # ---------------------------[ Import ]---------------------------

    try {
        Invoke-MgGraphRequest -Method POST -Uri $graphUri -Body ($policy | ConvertTo-Json -Depth 10 -Compress)
        Write-Log "Successfully imported: $($policy.displayName)" -Tag "Success"
    } catch {
        Write-Log "Failed to import $($file.Name): $_" -Tag "Error"
    }
}

# ---------------------------[ Script End ]---------------------------

Complete-Script -ExitCode 0