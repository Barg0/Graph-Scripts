# Script version:   2025-06-15 10:45
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$graphScopes = @(
    "Policy.Read.All",
    "Policy.ReadWrite.MobilityManagement",
    "DeviceManagementConfiguration.ReadWrite.All",
    "DeviceManagementServiceConfig.ReadWrite.All"
)

$scriptName = "Set-IntuneDefaultSettings"
$logFileName = "$($scriptName).log"

$logFileDirectory = $PSScriptRoot
$logFile = Join-Path -Path $logFileDirectory -ChildPath $logFileName

$log = $true
$enableLogFile = $false

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
        "Debug"   { "DarkYellow" }
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
    if ($null -ne $context.Account -and $null -ne $context.Scopes -and ($context.Scopes | Where-Object { $graphScopes -contains $_ })) {
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

# ---------------------------[ MDM user scope ]---------------------------

function Get-MdmUserScopePolicyId {
    $uri = "https://graph.microsoft.com/beta/policies/mobileDeviceManagementPolicies"

    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $mdmUserScopePolicy = $response.value | Where-Object { $_.displayName -eq "Microsoft Intune" }

        if (-not $mdmUserScopePolicy) {
            Write-Log "MDM user scope setting not found." -Tag "Error"
            return $null
        }

        $settingId = $mdmUserScopePolicy.id
        # Write-Log "Found MDM user scope setting ID: $settingId" -Tag "Debug"
        return $settingId
    } catch {
        Write-Log "Failed to retrieve MDM user scope settings: $_" -Tag "Error"
        return $null
    }
}

function Set-MdmUserScope {
    $uri = "https://graph.microsoft.com/beta/policies/mobileDeviceManagementPolicies/$mdmUserScopePolicyId"

    $body = @{
        appliesTo = "all"
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "MDM user scope configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update MDM user scope configuration: $_" -Tag "Error"
    }
}

# ---------------------------[ Enrollment restrictions ]---------------------------

function Get-EnrollmentLimitPolicyId {
    $uri = "https://graph.microsoft.com/beta//deviceManagement/deviceEnrollmentConfigurations"

    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $defaultEnrollmentLimitPolicy = $response.value | Where-Object { ($_.id.EndsWith("_DefaultLimit")) -and ($_.deviceEnrollmentConfigurationType -eq "limit" ) }

        if (-not $defaultEnrollmentLimitPolicy) {
            Write-Log "Default Enrollment Limit Policy not found." -Tag "Error"
            return $null
        }

        $settingId = $defaultEnrollmentLimitPolicy.id
        # Write-Log "Found Default Enrollment Limit Policy Id: $settingId" -Tag "Debug"
        return $settingId
    } catch {
        Write-Log "Failed to retrieve Default Enrollment Limit Policy: $_" -Tag "Error"
        return $null
    }
}

function Set-EnrollmentLimit {
    $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations/$enrollmentLimitPolicyId"

    $body = @{
        "@odata.type" = "#microsoft.graph.deviceEnrollmentLimitConfiguration"
        limit = "15" 
    }
    # Default Limit is 15
    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Enrollment Limit updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to Enrollment Limit configuration: $_" -Tag "Error"
    }
}

function Get-EnrollmentPlatformRestrictionsPolicyId {
    $uri = "https://graph.microsoft.com/beta//deviceManagement/deviceEnrollmentConfigurations"

    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $defaultEnrollmentPlatformRestrictionsPolicy = $response.value | Where-Object { ($_.id.EndsWith("_DefaultPlatformRestrictions")) -and ($_.deviceEnrollmentConfigurationType -eq "platformRestrictions") }

        if (-not $defaultEnrollmentPlatformRestrictionsPolicy) {
            Write-Log "Default Enrollment Platform Restrictions Policy not found." -Tag "Error"
            return $null
        }

        $settingId = $defaultEnrollmentPlatformRestrictionsPolicy.id
        # Write-Log "Found Enrollment Platform Restrictions Policy Id: $settingId" -Tag "Debug"
        return $settingId
    } catch {
        Write-Log "Failed to retrieve Enrollment Platform Restrictions Policy Policy: $_" -Tag "Error"
        return $null
    }
}

function Set-EnrollmentPlatformRestrictions {
    $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations/$enrollmentPlatformRestrictionsPolicyId"

    $body = @{
        "@odata.type" = "#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration"
        windowsRestriction = @{
            platformBlocked = $false
            personalDeviceEnrollmentBlocked = $true
        }
        windowsHomeSkuRestriction = @{
            platformBlocked = $false
            personalDeviceEnrollmentBlocked = $true
        }        
        macOSRestriction = @{
            platformBlocked = $false
            personalDeviceEnrollmentBlocked = $true
        }
        androidRestriction = @{
            platformBlocked = $true
            personalDeviceEnrollmentBlocked = $true
        }
        androidForWorkRestriction = @{
            platformBlocked = $false
            personalDeviceEnrollmentBlocked = $true
        }        
        iosRestriction = @{
            platformBlocked = $false
            personalDeviceEnrollmentBlocked = $true
        }
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Enrollment Platform Restrictions updated successfully." -Tag "Success"
    } catch {
        Write-Log "Enrollment Platform Restrictions configuration: $_" -Tag "Error"
    }
}

# ---------------------------[ Windows Hello for Business ]---------------------------

function Get-WindowsHelloForBusinessPolicyId {
    $uri = "https://graph.microsoft.com/beta//deviceManagement/deviceEnrollmentConfigurations"

    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $defaultWindowsHelloForBusinessPolicy = $response.value | Where-Object { ($_.id.EndsWith("_DefaultWindowsHelloForBusiness")) -and ($_.deviceEnrollmentConfigurationType -eq "windowsHelloForBusiness") }

        if (-not $defaultWindowsHelloForBusinessPolicy) {
            Write-Log "Default Windows Hello For Business Policy not found." -Tag "Error"
            return $null
        }

        $settingId = $defaultWindowsHelloForBusinessPolicy.id
        # Write-Log "Found Enrollment Platform Restrictions Policy Id: $settingId" -Tag "Debug"
        return $settingId
    } catch {
        Write-Log "Failed to retrieve Windows Hello For Business Policy Policy: $_" -Tag "Error"
        return $null
    }
}

function Set-WindowsHelloForBusiness {
    $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations/$windowsHelloForBusinessPolicyId"

    $body = @{
        "@odata.type" = "#microsoft.graph.deviceEnrollmentWindowsHelloForBusinessConfiguration"
        state = "disabled"
        securityKeyForSignIn = "notConfigured"
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Windows Hello For Business configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Windows Hello For Business configuration: $_" -Tag "Error"
    }
}

# ---------------------------[ Enrollment Status Page ]---------------------------

function Get-EnrollmentStatusPagePolicyId {
    $uri = "https://graph.microsoft.com/beta//deviceManagement/deviceEnrollmentConfigurations"

    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $defaultWindowsHelloForBusinessPolicy = $response.value | Where-Object { ($_.id.EndsWith("_DefaultWindows10EnrollmentCompletionPageConfiguration")) -and ($_.deviceEnrollmentConfigurationType -eq "windows10EnrollmentCompletionPageConfiguration") }

        if (-not $defaultWindowsHelloForBusinessPolicy) {
            Write-Log "Default Enrollment Status Page Policy not found." -Tag "Error"
            return $null
        }

        $settingId = $defaultWindowsHelloForBusinessPolicy.id
        # Write-Log "Found Enrollment Status Page Id: $settingId" -Tag "Debug"
        return $settingId
    } catch {
        Write-Log "Failed to retrieve Enrollment Status Page Policy Policy: $_" -Tag "Error"
        return $null
    }
}

function Set-EnrollmentStatusPage {
    $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations/$enrollmentStatusPagePolicyId"

    $body = @{
        "@odata.type" = "#microsoft.graph.windows10EnrollmentCompletionPageConfiguration"
        showInstallationProgress = $true
        blockDeviceSetupRetryByUser = $false
        allowDeviceResetOnInstallFailure = $true
        allowLogCollectionOnInstallFailure = $true
        customErrorMessage = ""
        installProgressTimeoutInMinutes = "60"
        allowDeviceUseOnInstallFailure = $false
        selectedMobileAppIds = @()
        allowNonBlockingAppInstallation = $false
        installQualityUpdates = $false
        trackInstallProgressForAutopilotOnly = $false
        disableUserStatusTrackingAfterFirstUser = $false
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Enrollment Status Page configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Enrollment Status Page configuration: $_" -Tag "Error"
    }
}

# ---------------------------[ Execution ]---------------------------

$mdmUserScopePolicyId = Get-MdmUserScopePolicyId
Set-MdmUserScope
$enrollmentLimitPolicyId = Get-EnrollmentLimitPolicyId
Set-EnrollmentLimit
$enrollmentPlatformRestrictionsPolicyId = Get-EnrollmentPlatformRestrictionsPolicyId
Set-EnrollmentPlatformRestrictions
$windowsHelloForBusinessPolicyId = Get-WindowsHelloForBusinessPolicyId
Set-WindowsHelloForBusiness
$enrollmentStatusPagePolicyId = Get-EnrollmentStatusPagePolicyId
Set-EnrollmentStatusPage

# ---------------------------[ Script Complete ]---------------------------

Complete-Script -ExitCode 0