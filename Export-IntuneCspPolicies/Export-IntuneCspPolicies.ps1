# ====================[ Script Metadata ]====================
# Script version:   2025-06-01 17:45
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

$scriptStartTime = Get-Date

# ---------------------------[ Script name ]---------------------------

$scriptName = "Export-IntuneCspPolicies"
$logFileName = "$($scriptName).log"

$exportPath = Join-Path -Path $PSScriptRoot -ChildPath "Export"

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

# ---------------------------[ Graph Connection ]---------------------------

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
    $context = Get-MgContext
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

# ---------------------------[ Create Export Folder ]---------------------------

try {
    New-Item -ItemType Directory -Force -Path $exportPath | Out-Null
    Write-Log "Created export folder at $exportPath" -Tag "Info"
} catch {
    Write-Log "Failed to create export directory: $_" -Tag "Error"
    Complete-Script -ExitCode 2
}

# ---------------------------[ Get Custom Profiles ]---------------------------
try {
    $allConfigs = Get-MgDeviceManagementDeviceConfiguration
    $customConfigs = $allProfiles | Where-Object {
    ($_.AdditionalProperties.'@odata.type' -notlike "#microsoft.graph.macOS*") -and
    ($_.AdditionalProperties.'@odata.type' -notlike "#microsoft.graph.ios*") -and
    ($_.AdditionalProperties.'@odata.type' -notlike "#microsoft.graph.android*") -and
    ($_.AdditionalProperties.'@odata.type' -ne "#microsoft.graph.windows10CustomConfiguration")
}


    if ($customConfigs.Count -eq 0) {
        Write-Log "No custom OMA-URI configuration profiles found." -Tag "Error"
        Complete-Script -ExitCode 3
    } else {
        Write-Log "$($customConfigs.Count) custom configuration profiles found." -Tag "Info"
    }
} catch {
    Write-Log "Error retrieving device configurations: $_" -Tag "Error"
    Complete-Script -ExitCode 4
}

# ---------------------------[ Export Profiles to JSON ]---------------------------

foreach ($config in $customConfigs) {
    try {
        Write-Log "Processing profile: $($config.DisplayName)" -Tag "Info"
        $fullConfig = Get-MgDeviceManagementDeviceConfiguration -DeviceConfigurationId $config.Id

        $exportObject = [PSCustomObject]@{
            displayName = $fullConfig.DisplayName
            description = $fullConfig.Description
            omaSettings = @()
        }

        if ($fullConfig.AdditionalProperties.ContainsKey("omaSettings")) {
            $omaList = $fullConfig.AdditionalProperties["omaSettings"]
            foreach ($setting in $omaList) {
                $exportObject.omaSettings += [PSCustomObject]@{
                    displayName = $setting.displayName
                    description = $setting.description
                    omaUri     = $setting.omaUri
                    value      = $setting.value
                    odataType  = $setting.'@odata.type'
                }
            }

            $fileName = ($fullConfig.DisplayName -replace '[\\\/:*?"<>|]', '_') + ".json"
            $filePath = Join-Path $exportPath $fileName

            $exportObject | ConvertTo-Json -Depth 10 | Out-File -Encoding UTF8 $filePath
            Write-Log "Exported profile to $filePath" -Tag "Success"
        } else {
            Write-Log "Profile '$($config.DisplayName)' has no omaSettings. Skipping export." -Tag "Info"
        }

    } catch {
        Write-Log "Failed to export profile '$($config.DisplayName)': $_" -Tag "Error"
    }
}

# ---------------------------[ Script End ]---------------------------

Complete-Script -ExitCode 0
