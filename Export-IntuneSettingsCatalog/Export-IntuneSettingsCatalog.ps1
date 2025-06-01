# Script version:   2025-06-01 19:40
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$exportPath = $exportPath = Join-Path -Path $PSScriptRoot -ChildPath "Export"
$endpointUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
$GraphScope = "DeviceManagementConfiguration.Read.All"

# ---------------------------[ Script name ]---------------------------

$scriptName = "Export-IntuneSettingsCatalog"
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

# ---------------------------[ Create Export Folder ]---------------------------

try {
    New-Item -ItemType Directory -Force -Path $exportPath | Out-Null
    Write-Log "Created export folder at $exportPath" -Tag "Info"
} catch {
    Write-Log "Failed to create export directory: $_" -Tag "Error"
    Complete-Script -ExitCode 2
}

# ---------------------------[ Connect to Microsoft Graph ]---------------------------

# Ensure the Microsoft.Graph SDK is installed
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

# Import the module
try {
    Import-Module Microsoft.Graph -Force -ErrorAction Stop
    Write-Log "Microsoft.Graph module imported." -Tag "Success"
} catch {
    Write-Log "Failed to import Microsoft.Graph module: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

# Validate Graph connection and scope
$connected = $false
try {
    $context = Get-MgContext $GraphScope
    if ($null -ne $context.Account -and $null -ne $context.Scopes -and $context.Scopes -contains "DeviceManagementConfiguration.Read.All") {
        Write-Log "Microsoft Graph already connected as $($context.Account)" -Tag "Success"
        $connected = $true
    } else {
        Write-Log "Microsoft Graph context incomplete or lacks required scope. Reconnecting..." -Tag "Info"
    }
} catch {
    Write-Log "Microsoft Graph not connected. Attempting connection..." -Tag "Info"
}

# Connect if not already valid
if (-not $connected) {
    try {
        Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All" | Out-Null
        Write-Log "Connected to Microsoft Graph successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

# ---------------------------[ Query Configuration Policies ]---------------------------
function Get-GraphDataRecursive {
    param ([string]$InitialUri)
    $results = @()
    $currentUri = $InitialUri

    while ($null -ne $currentUri -and $currentUri -ne '') {
        Write-Log "Querying: $currentUri" -Tag "Debug"
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $currentUri
            $results += $response.value
            $currentUri = $response.'@odata.nextLink'
        } catch {
            Write-Log "Failed to retrieve: $currentUri - $_" -Tag "Error"
            break
        }
    }

    return $results
}

Write-Log "Fetching Settings Catalog policies..." -Tag "Check"
$policies = Get-GraphDataRecursive -InitialUri $endpointUri
Write-Log "Retrieved $($policies.Count) policies" -Tag "Success"

# ---------------------------[ Export to JSON ]---------------------------

foreach ($policy in $policies) {
    $policyId = $policy.id

    # Use displayName for filename if present, else fallback to ID
    $safeName = if ($null -ne $policy.displayName -and $policy.displayName -ne '') {
        ($policy.displayName -replace '[\\/:*?"<>|]', '_')
    } else {
        $policyId
    }

    $filePath = Join-Path -Path $ExportPath -ChildPath "$($policy.name).json"

    try {
        # Fetch settings separately
        $settingsUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$policyId/settings"
        $settingsResponse = Invoke-MgGraphRequest -Method GET -Uri $settingsUri

        # Build full export object
        $fullExport = @{
            id           = $policy.id
            displayName  = $policy.name
            description  = $policy.description
            platforms    = $policy.platforms
            technologies = $policy.technologies
            settings     = $settingsResponse.value
        }

        # Save combined data to JSON file
        $fullExport | ConvertTo-Json -Depth 20 | Out-File -FilePath $filePath -Encoding utf8
        Write-Log "Exported: $($policy.name) ($safeName)" -Tag "Success"
    } catch {
        Write-Log "Failed to export policy ID $($policyId): $_" -Tag "Error"
    }
}

# ---------------------------[ End ]---------------------------

Complete-Script -ExitCode 0
