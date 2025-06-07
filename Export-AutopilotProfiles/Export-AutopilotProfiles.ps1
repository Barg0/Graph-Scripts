# Script version:   2025-06-07 17:15
# Script author:    Barg0

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$graphScopes = @(
    "DeviceManagementServiceConfig.Read.All"
)
$uri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles"

$scriptName = "Export-AutopilotProfiles"
$logFileName = "$($scriptName).log"

# Output folder for JSON
$exportDir = Join-Path -Path $PSScriptRoot -ChildPath "ExportAutopilotProfiles"

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

function Write-Log {
    param ([string]$Message, [string]$Tag = "Info")

    if (-not $log) { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $tagList = @("Start", "Check", "Info", "Success", "Error", "Debug", "End")
    $rawTag = $Tag.Trim()

    if ($tagList -contains $rawTag) {
        $rawTag = $rawTag.PadRight(7)
    } else {
        $rawTag = "Error  "
    }

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

    if ($enableLogFile) {
        "$logMessage" | Out-File -FilePath $logFile -Append
    }

    Write-Host "$timestamp " -NoNewline
    Write-Host "[  " -NoNewline -ForegroundColor White
    Write-Host "$rawTag" -NoNewline -ForegroundColor $color
    Write-Host " ] " -NoNewline -ForegroundColor White
    Write-Host "$Message"
}

function Complete-Script {
    param([int]$ExitCode)
    $scriptEndTime = Get-Date
    $duration = $scriptEndTime - $scriptStartTime
    Write-Log "Script execution time: $($duration.ToString("hh\:mm\:ss\.ff"))" -Tag "Info"
    Write-Log "Exit Code: $ExitCode" -Tag "Info"
    Write-Log "======== Script Completed ========" -Tag "End"
    exit $ExitCode
}

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

# ---------------------------[ Create Export Folder ]---------------------------

if (-not (Test-Path -Path $exportDir)) {
    try {
        New-Item -ItemType Directory -Force -Path $exportDir | Out-Null
        Write-Log "Created export folder at $exportDir" -Tag "Info"
    } catch {
        Write-Log "Failed to create export directory at $($exportDir): $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

# ---------------------------[ Graph API Request ]---------------------------

Write-Log "Graph URI: $uri" -Tag "Info"

try {
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
} catch {
    Write-Log "Graph API request failed: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

# ---------------------------[ Filter Autopilot Profiles ]---------------------------

try {
    $profiles = $response.value

    if ($null -eq $profiles -or $profiles.Count -eq 0) {
        Write-Log "No Autopilot deployment profiles found." -Tag "Info"
        Complete-Script -ExitCode 0
    } else {
        Write-Log "Found $($profiles.Count) Autopilot deployment profiles." -Tag "Info"
    }
} catch {
    Write-Log "Failed to filter Autopilot profiles: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

# ---------------------------[ Export Autopilot Profiles ]---------------------------

foreach ($profile in $profiles) {

    $profileName = $profile.displayName -replace '[\\/:*?"<>|]', '_'
    $fileName = "$profileName.json"
    $exportPath = Join-Path -Path $exportDir -ChildPath $fileName

    # Prepare full profile object
    $fullProfile = $profile.PSObject.Copy()

    try {
        $fullProfile | ConvertTo-Json -Depth 20 | Out-File -FilePath $exportPath -Encoding UTF8
        Write-Log "Exported: $fileName" -Tag "Success"
    } catch {
        Write-Log "Failed to export $($profile.displayName): $_" -Tag "Error"
    }
}

# ---------------------------[ Script Completion ]---------------------------

Complete-Script -ExitCode 0