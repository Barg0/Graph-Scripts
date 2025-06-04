# Script version:   2025-06-04 13:30
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

# Capture start time to log script duration
$scriptStartTime = Get-Date

# ---------------------------[ Paramter ]---------------------------

$GraphScope = "Policy.Read.All"
$graphUri = "https://graph.microsoft.com/beta/identity/conditionalAccess/policies"
$outputDir = Join-Path -Path $PSScriptRoot -ChildPath "Export"

# ---------------------------[ Script name ]---------------------------

# Script name used for folder/log naming
$scriptName = "Export-ConditionalAccessPolicies"
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
    if ($null -ne $context.Account -and $null -ne $context.Scopes -and $context.Scopes -contains $GraphScope) {
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
        Connect-MgGraph -Scopes $GraphScope | Out-Null
        Write-Log "Connected to Microsoft Graph successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

# ---------------------------[ Graph API Request ]---------------------------

$uri = $graphUri
Write-Log "Calling Graph API for Conditional Access policies..." -Tag "Info"

try {
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    Write-Log "Successfully retrieved Conditional Access policies." -Tag "Success"
} catch {
    Write-Log "Graph API request failed: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

# ---------------------------[ Create Export Folder ]---------------------------

if (-not (Test-Path -Path $outputDir)) {
    try {
        New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
        Write-Log "Created export folder at $outputDir" -Tag "Info"
    } catch {
        Write-Log "Failed to create export directory at $($outputDir): $_" -Tag "Error"
        Complete-Script -ExitCode 2
    }
} else {
    Write-Log "Export folder already exists at $outputDir" -Tag "Info"
}

# ---------------------------[ Export Policies ]---------------------------

$policyCount = $response.value.Count
Write-Log "Exporting $policyCount policies to individual JSON files..." -Tag "Info"

foreach ($policy in $response.value) {
    # Use displayName or fallback to ID
    $name = if ($policy.displayName) { $policy.displayName } else { $policy.id }

    # Sanitize filename
    $safeName = ($name -replace '[\\/:*?"<>|]', '_') -replace '\s+', '_'
    $fileName = "$safeName.json"
    $filePath = Join-Path -Path $outputDir -ChildPath $fileName

    try {
        $policy | ConvertTo-Json -Depth 10 | Out-File -FilePath $filePath -Encoding UTF8
        Write-Log "Exported: $fileName" -Tag "Success"
    } catch {
        Write-Log "Failed to export policy '$name': $_" -Tag "Error"
    }
}

Write-Log "All policies exported to: $outputDir" -Tag "Success"

# ---------------------------[ Script End ]---------------------------

Complete-Script -ExitCode 0