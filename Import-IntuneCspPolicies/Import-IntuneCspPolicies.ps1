# Script version:   2025-06-07 17:15
# Script author:    Barg0

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$graphScopes = @(
    "DeviceManagementConfiguration.Read.All"
    "DeviceManagementConfiguration.ReadWrite.All",
    "Directory.Read.All"
)
$uri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"

$scriptName = "Import-IntuneCspPolicies"
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
}

# ---------------------------[ Graph Authentication ]---------------------------

$connected = $false
try {
    $context = Get-MgContext
    if ($null -ne $context.Account -and $null -ne $context.Scopes -and $graphScopes | ForEach-Object { $context.Scopes -contains $_ }) {
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

# ---------------------------[ Get Tenant ID ]---------------------------

try {
    $tenantResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction Stop
    $tenantId = $tenantResponse.value[0].id
    Write-Log "Retrieved tenant ID: $tenantId" -Tag "Info"
} catch {
    Write-Log "Failed to retrieve tenant ID: $_" -Tag "Error"
    Complete-Script -ExitCode 2
}

# ---------------------------[ Fetch Existing DisplayNames ]---------------------------

try {
    $existingResponse = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    $existingConfigs = $existingResponse.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.windows10CustomConfiguration' }
    $existingNames = $existingConfigs.displayName
    Write-Log "Retrieved $($existingNames.Count) existing custom policies for comparison." -Tag "Info"
} catch {
    Write-Log "Failed to fetch existing configurations: $_" -Tag "Error"
    Complete-Script -ExitCode 5
}

# ---------------------------[ Import JSON Files ]---------------------------

if (-not (Test-Path -Path $importDir)) {
    Write-Log "Import directory not found: $importDir" -Tag "Error"
    Complete-Script -ExitCode 1
}

$files = Get-ChildItem -Path $importDir -Filter *.json
if ($files.Count -eq 0) {
    Write-Log "No JSON files found in $importDir" -Tag "Error"
    Complete-Script -ExitCode 1
}

foreach ($file in $files) {
    try {
        $json = Get-Content -Raw -Path $file.FullName | ConvertFrom-Json
        $profileName = $json.displayName

        if ($existingNames -contains $profileName) {
            Write-Log "Policy '$profileName' already exists. Skipping." -Tag "Info"
            continue
        }

        # Build properly typed omaSettings array
        $omaSettings = @()
        foreach ($s in $json.omaSettings) {
            $omaSettings += [pscustomobject]@{
                '@odata.type' = $s.odataType
                displayName   = $s.displayName
                description   = $s.description
                omaUri        = ($s.omaUri -replace '\{TenantID\}', $tenantId)
                value         = $s.value
            }
        }

        # Compose payload with proper formatting
        $payload = [pscustomobject]@{
            '@odata.type' = '#microsoft.graph.windows10CustomConfiguration'
            displayName   = $json.displayName
            description   = $json.description
            omaSettings   = $omaSettings
        }

        $body = $payload | ConvertTo-Json -Depth 10 -Compress
        # Write-Log "Importing profile: $profileName" -Tag "Debug"

        Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
        Write-Log "Successfully imported '$profileName'" -Tag "Success"
    } catch {
        Write-Log "Failed to import '$($file.Name)': $_" -Tag "Error"
    }
}

Complete-Script -ExitCode 0
