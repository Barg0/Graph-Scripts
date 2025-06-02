# Script version:   2025-05-29 11:10
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$importPath = Join-Path -Path $PSScriptRoot -ChildPath "Import"

# ---------------------------[ Script name ]---------------------------

$scriptName = "Import-IntuneSettingsCatalog"
$logFileName = "$($scriptName).log" 

# ---------------------------[ Logging Setup ]---------------------------

# Logging control switches
$log = $true                     # Set to $false to disable logging in shell
$enableLogFile = $false          # Set to $false to disable file output

# Define the log output location
$logFileDirectory = $PSScriptRoot
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
    if ($null -ne $context.Account -and $null -ne $context.Scopes -and $context.Scopes -contains "DeviceManagementConfiguration.ReadWrite.All") {
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
        Connect-MgGraph -Scopes "DeviceManagementConfiguration.ReadWrite.All" | Out-Null
        Write-Log "Connected to Microsoft Graph successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

# ---------------------------[ Import ]---------------------------

function Import-IntuneProfile {
    param (
        [string]$FilePath
    )

    Write-Log "Processing file: $FilePath" -Tag "Info"

    try {
        $raw = Get-Content -Path $FilePath -Raw | ConvertFrom-Json

        $profile = [ordered]@{
            name = $raw.name
            description = $raw.description
            platforms = $raw.platforms
            technologies = $raw.technologies
            roleScopeTagIds = $raw.roleScopeTagIds
            settings = $raw.settings
        }

        $body = $profile | ConvertTo-Json -Depth 20 -Compress

        $response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body -ContentType 'application/json'

        if ($null -ne $response.id) {
            Write-Log "Imported profile '$($response.name)' successfully with ID $($response.id)" -Tag "Success"
        } else {
            Write-Log "No ID returned. Import may have failed." -Tag "Warn"
        }
    }
    catch {
        Write-Log "Failed to import '$FilePath': $_" -Tag "Error"
    }
}

Get-ChildItem -Path $importPath -Filter *.json | ForEach-Object {
    Import-IntuneProfile -FilePath $_.FullName
}

Write-Log "All profiles processed." -Tag "End"
# Disconnect-MgGraph
Complete-Script -ExitCode 0
