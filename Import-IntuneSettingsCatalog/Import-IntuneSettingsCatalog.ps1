# Script version:   2025-06-01 20:40
# Script author:    Barg0

# ---------------------------[ Description ]---------------------------

# Please run this before you start the script: 
# Disconnect-MgGraph
# Connect-MgGraph -Scopes "DeviceManagementConfiguration.ReadWrite.All"

# ---------------------------[ Script Start Timestamp ]---------------------------

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$importPath = Join-Path -Path $PSScriptRoot -ChildPath "Import"
$GraphScope = "DeviceManagementConfiguration.Read.All"

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

# ---------------------------[ Import ]---------------------------

Get-ChildItem -Path $ImportPath -Filter *.json | ForEach-Object {
    try {
        $json = Get-Content $_.FullName -Raw | ConvertFrom-Json
        $sanitizedSettings = @()

        foreach ($setting in $json.settings) {
            $instance = $setting.settingInstance
            $sanitized = [ordered]@{
                settingDefinitionId = $instance.settingDefinitionId
                "@odata.type"       = $instance.'@odata.type'
            }

            if ($instance.simpleSettingValue) {
                $sanitized.simpleSettingValue = @{
                    value = $instance.simpleSettingValue.value
                }
            }
            elseif ($instance.choiceSettingValue) {
                $sanitized.choiceSettingValue = @{
                    value    = $instance.choiceSettingValue.value
                    children = $instance.choiceSettingValue.children
                }
            }

            $sanitizedSettings += @{
                settingInstance = $sanitized
            }
        }

        $body = @{
            description  = $json.description
            name         = $json.displayName
            technologies = $json.technologies
            platforms    = $json.platforms
            settings     = $sanitizedSettings
        }

        $response = Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" `
            -Body ($body | ConvertTo-Json -Depth 20) `
            -ContentType 'application/json'

        Write-Log "Imported: $($_.Name)" -Tag "Success"
    }
    catch {
        Write-Log "Error importing $($_.Name): $($_.Exception.Message)" -Tag "Error"
    }
}

Complete-Script -ExitCode 0