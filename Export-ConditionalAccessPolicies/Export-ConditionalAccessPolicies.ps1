# Script version:   2025-06-08 10:00
# Script author:    Barg0

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$graphScopes = @(
    "Policy.Read.All"
)
$uri = "https://graph.microsoft.com/beta/identity/conditionalAccess/policies"

$scriptName = "Export-ConditionalAccessPolicies"
$logFileName = "$($scriptName).log"

# Output folder for JSON
$exportDir = Join-Path -Path $PSScriptRoot -ChildPath "ExportConditionalAccess"

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
    if ($null -ne $context.Account -and $null -ne $context.Scopes -and $graphScopes | ForEach-Object { $context.Scopes -contains $_ }) {
        Write-Log "Microsoft Graph already connected as $($context.Account)" -Tag "Success"
        $connected = $true
    } else {
        Write-Log "Microsoft Graph context incomplete or lacks required scopes. Reconnecting..." -Tag "Info"
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

# ---------------------------[ Create Export Folder ]---------------------------

if (-not (Test-Path -Path $exportDir)) {
    try {
        New-Item -ItemType Directory -Force -Path $exportDir | Out-Null
        Write-Log "Created export folder at $exportDir" -Tag "Info"
    } catch {
        Write-Log "Failed to create export directory: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

# ---------------------------[ Graph API Request ]---------------------------

Write-Log "Querying Conditional Access policies from Microsoft Graph..." -Tag "Info"
Write-Log "Graph URI: $uri" -Tag "Debug"

try {
    $policies = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    if ($policies.value.Count -eq 0) {
        Write-Log "No Conditional Access policies found." -Tag "Info"
    } else {
        Write-Log "Found $($policies.value.Count) Conditional Access policies." -Tag "Success"

        foreach ($policy in $policies.value) {
            $fileName = ($policy.displayName -replace '[\\/:*?"<>|]', "_") + ".json"
            $exportPath = Join-Path -Path $exportDir -ChildPath $fileName
            try {
                $policy | ConvertTo-Json -Depth 10 | Out-File -FilePath $exportPath -Encoding UTF8
                Write-Log "Exported policy '$($policy.displayName)' to '$fileName'" -Tag "Success"
            } catch {
                Write-Log "Failed to export policy '$($policy.displayName)': $_" -Tag "Error"
            }
        }
    }
} catch {
    Write-Log "Graph API request failed: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

Complete-Script -ExitCode 0
