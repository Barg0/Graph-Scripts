# Script version:   2025-06-07 13:15
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$graphScopes = @(
    "DeviceManagementConfiguration.ReadWrite.All"
)
$uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"

$scriptName = "Import-IntuneSecurityPolicies"
$logFileName = "$($scriptName).log"

# Import folder for JSON
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
} else {
    Write-Log "Microsoft.Graph module is already loaded." -Tag "Info"
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

# ---------------------------[ Import Policies ]---------------------------

if (-not (Test-Path -Path $importDir)) {
    Write-Log "Import directory not found: $importDir" -Tag "Error"
    Complete-Script -ExitCode 1
}

# Get all existing policies to check for name collisions
$existingPolicies = Invoke-MgGraphRequest -Method GET -Uri $uri

Get-ChildItem -Path $importDir -Filter *.json | ForEach-Object {
    try {
        $jsonPath = $_.FullName
        $json = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json -Depth 20
        $policyName = $json.name

        # Check for name duplication
        $match = $existingPolicies.value | Where-Object { $_.name -eq $policyName }
        if ($null -ne $match) {
            Write-Log "Policy '$policyName' already exists. Skipping import." -Tag "Info"
            return
        }

        # Construct POST body with required fields including 'settings'
        $bodyObj = [ordered]@{
            name              = $json.name
            description       = $json.description
            platforms         = $json.platforms
            technologies      = $json.technologies
            roleScopeTagIds   = $json.roleScopeTagIds
            templateReference = $json.templateReference
            settings          = $json.settings
        }

        $body = $bodyObj | ConvertTo-Json -Depth 20

        # POST to create the policy with settings
        # $createUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
        $newPolicy = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json"
        $newId = $newPolicy.id
        Write-Log "Created policy '$policyName' with ID $newId" -Tag "Success"

    } catch {
        Write-Log "Unexpected error importing policy file '$($_.Name)': $_" -Tag "Error"
    }
}


Complete-Script -ExitCode 0


# $settingsUri = "$($uri)/$($newId)/settings"