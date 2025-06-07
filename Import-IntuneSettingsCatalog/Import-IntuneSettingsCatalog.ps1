# Script version:   2025-06-07 13:45
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$graphScopes = @(
    "DeviceManagementConfiguration.ReadWrite.All"
)
$uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
$scriptName = "Import-IntuneSettingsCatalog"
$logFileName = "$scriptName.log"

# Import folder for JSON
$importDir = Join-Path -Path $PSScriptRoot -ChildPath "Import"

# ---------------------------[ Logging Control ]---------------------------

$log = $true                     # Set to $false to disable logging in shell
$enableLogFile = $false          # Set to $false to disable file output
$logFileDirectory = "$PSScriptRoot"
$logFile = Join-Path -Path $logFileDirectory -ChildPath $logFileName

# ---------------------------[ Logging Setup ]---------------------------

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
} else {
    Write-Log "Microsoft.Graph module is already loaded." -Tag "Info"
}

# ---------------------------[ Graph Authentication ]---------------------------

$connected = $false
try {
    $context = Get-MgContext
    if ($null -ne $context.Account -and $null -ne $context.Scopes -and ($graphScopes | ForEach-Object { $_ -in $context.Scopes })) {
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

# ---------------------------[ Import Function ]---------------------------

function Import-IntuneProfile {
    param (
        [string]$FilePath,
        [string]$GraphUri,
        [hashtable]$ExistingProfiles
    )

    # Write-Log "Processing file: $FilePath" -Tag "Debug"

    try {
        $rawContent = Get-Content -Path $FilePath -Raw
        $profileData = $rawContent | ConvertFrom-Json -ErrorAction Stop

        if (-not $profileData.name) {
            Write-Log "Profile in '$FilePath' is missing a 'name' property. Skipping." -Tag "Error"
            return
        }

        $profileName = $profileData.name

        if ($ExistingProfiles.ContainsKey($profileName)) {
            Write-Log "Profile '$profileName' already exists. Skipping import." -Tag "Info"
            return
        }

        $body = @{
            name             = $profileData.name
            description      = $profileData.description
            platforms        = $profileData.platforms
            technologies     = $profileData.technologies
            roleScopeTagIds  = $profileData.roleScopeTagIds
            settings         = $profileData.settings
        } | ConvertTo-Json -Depth 20 -Compress

        $response = Invoke-MgGraphRequest -Method POST -Uri $GraphUri -Body $body -ContentType 'application/json'

        if ($null -ne $response.id) {
            Write-Log "Imported profile '$($response.name)' successfully with ID $($response.id)" -Tag "Success"
        } else {
            Write-Log "No ID returned for '$profileName'. Import may have failed." -Tag "Info"
        }
    } catch {
        Write-Log "Exception while importing '$FilePath': $($_.Exception.Message)" -Tag "Error"
    }
}

# ---------------------------[ Import Profiles ]---------------------------

if (-not (Test-Path -Path $importDir)) {
    Write-Log "Import directory not found: $importDir" -Tag "Error"
    Complete-Script -ExitCode 1
}

Write-Log "Retrieving existing Intune configuration profiles..." -Tag "Info"

$existingProfiles = @{}
$nextLink = $uri

try {
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink
        foreach ($profile in $response.value) {
            $normalizedName = ($profile.name -replace '\s+', ' ').Trim()
            $existingProfiles[$normalizedName] = $profile.id
        }
        $nextLink = $response.'@odata.nextLink'
    } while ($null -ne $nextLink)

    Write-Log "Loaded $($existingProfiles.Count) existing profiles." -Tag "Info"
} catch {
    Write-Log "Failed to retrieve existing profiles: $($_.Exception.Message)" -Tag "Error"
    Complete-Script -ExitCode 1
}

Get-ChildItem -Path $importDir -Filter *.json | ForEach-Object {
    Import-IntuneProfile -FilePath $_.FullName -GraphUri $uri -ExistingProfiles $existingProfiles
}

Complete-Script -ExitCode 0