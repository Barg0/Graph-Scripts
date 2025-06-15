# Script version:   2025-06-15 10:45
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$graphScopes = @(
    "Policy.Read.All",
    "Policy.ReadWrite.ConditionalAccess",
    "Directory.Read.All",
    "User.Read.All",
    "Policy.ReadWrite.AuthenticationMethod",
    "Policy.ReadWrite.DeviceConfiguration",
    "Directory.ReadWrite.All",
    "Policy.ReadWrite.Authorization	"
)

$scriptName = "Set-EntraDefaultSettings"
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

# ---------------------------[ Device Settings ]---------------------------

function Set-DeviceSettings {
    $uri = "https://graph.microsoft.com/beta/policies/deviceRegistrationPolicy"
    $body = @{
        userDeviceQuota = 50
        multiFactorAuthConfiguration = "notRequired"
        azureADRegistration = @{
            isAdminConfigurable = $false
            allowedToRegister = @{
                "@odata.type" = "#microsoft.graph.allDeviceRegistrationMembership"
            }
        }
        azureADJoin = @{
            isAdminConfigurable = $true
            allowedToJoin = @{
                "@odata.type" = "#microsoft.graph.allDeviceRegistrationMembership"
            }
            localAdmins = @{
                enableGlobalAdmins = $false
                registeringUsers = @{
                    "@odata.type" = "#microsoft.graph.noDeviceRegistrationMembership"
                }
            }
        }
        localAdminPassword = @{
            isEnabled = $true
        }
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        # Write-Log "Updating device registration policy settings..." -Tag "Debug"
        Invoke-MgGraphRequest -Method PUT -Uri $uri -Body $json -ErrorAction Stop | Out-Null
        Write-Log "Device settings updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update device settings: $_" -Tag "Error"
    }
}

# ---------------------------[ Authentication Methods ]---------------------------

function Set-Fido2Configuration {
    $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/fido2"

    $body = @{
        "@odata.type" = "#microsoft.graph.fido2AuthenticationMethodConfiguration"
        id = "Fido2"
        state = "enabled"
        isSelfServiceRegistrationAllowed = $true
        isAttestationEnforced = $false
        includeTargets = @(@{
            targetType = "group"
            id = "all_users"
            isRegistrationRequired = $false
        })
        excludeTargets = @()
        keyRestrictions = @{
            isEnforced = $true
            enforcementType = "allow"
            aaGuids = @(
                "90a3ccdf-635c-4729-a248-9b709135078f",
                "de1e552d-db1d-4423-a619-566b625cdc84"
            )
        }
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "FIDO2 configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update FIDO2 configuration: $_" -Tag "Error"
    }
}

function Set-MicrosoftAuthenticatorConfiguration {
    $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/MicrosoftAuthenticator"

    $body = @{
        "@odata.type" = "#microsoft.graph.microsoftAuthenticatorAuthenticationMethodConfiguration"
        id = "MicrosoftAuthenticator"
        state = "enabled"
        isSoftwareOathEnabled = $false
        includeTargets = @(@{
            targetType = "group"
            id = "all_users"
            isRegistrationRequired = $false
            authenticationMode = "any"
        })
        featureSettings = @{
            displayAppInformationRequiredState = @{
                state = "enabled"
                includeTarget = @{ targetType = "group"; id = "all_users" }
                excludeTarget = @{ targetType = "group"; id = "00000000-0000-0000-0000-000000000000" }
            }
            displayLocationInformationRequiredState = @{
                state = "enabled"
                includeTarget = @{ targetType = "group"; id = "all_users" }
                excludeTarget = @{ targetType = "group"; id = "00000000-0000-0000-0000-000000000000" }
            }
            companionAppAllowedState = @{
                state = "default"
                includeTarget = @{ targetType = "group"; id = "all_users" }
                excludeTarget = @{ targetType = "group"; id = "00000000-0000-0000-0000-000000000000" }
            }
        }
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Microsoft Authenticator configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update Microsoft Authenticator configuration: $_" -Tag "Error"
    }
}

function Set-TemporaryAccessPassConfiguration {
    $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/temporaryAccessPass"

    $body = @{
        "@odata.type" = "#microsoft.graph.temporaryAccessPassAuthenticationMethodConfiguration"
        state = "enabled"
        includeTargets = @(@{
            targetType = "group"
            id = "all_users"
            isRegistrationRequired = $false
        })
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Temporary Access Pass configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update Temporary Access Pass configuration: $_" -Tag "Error"
    }
}

function Set-SoftwareOathConfiguration {
    $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/softwareOath"

    $body = @{
        "@odata.type" = "#microsoft.graph.softwareOathAuthenticationMethodConfiguration"
        state = "enabled"
        includeTargets = @(@{
            targetType = "group"
            id = "all_users"
            isRegistrationRequired = $false
        })
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Software OATH configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update Software OATH configuration: $_" -Tag "Error"
    }
}

function Set-HardwareOathConfiguration {
    $uri = "https://graph.microsoft.com/beta/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/hardwareOath"

    $body = @{
        "@odata.type" = "#microsoft.graph.hardwareOathAuthenticationMethodConfiguration"
        id = "HardwareOath"
        state = "enabled"
        includeTargets = @(@{
            targetType             = "group"
            id                     = "all_users"
            isRegistrationRequired = $false
        })
        excludeTargets = @()
    }

    $json = $body | ConvertTo-Json -Depth 5

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Hardware OATH configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update Hardware OATH configuration: $_" -Tag "Error"
    }
}

function Set-EmailConfiguration {
    $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/email"

    $body = @{
        "@odata.type" = "#microsoft.graph.emailAuthenticationMethodConfiguration"
        state = "disabled"
        includeTargets = @(@{
            targetType = "group"
            id = "all_users"
            isRegistrationRequired = $false
        })
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "E-Mail configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update E-Mail configuration: $_" -Tag "Error"
    }
}

function Set-SmsConfiguration {
    $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/sms"

    $body = @{
        "@odata.type" = "#microsoft.graph.smsAuthenticationMethodConfiguration"
        state = "disabled"
        includeTargets = @(@{
            targetType = "group"
            id = "all_users"
            isRegistrationRequired = $false
        })
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "SMS configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update SMS configuration: $_" -Tag "Error"
    }
}

function Set-VoiceConfiguration {
    $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/voice"

    $body = @{
        "@odata.type" = "#microsoft.graph.voiceAuthenticationMethodConfiguration"
        state = "disabled"
        includeTargets = @(@{
            targetType = "group"
            id = "all_users"
            isRegistrationRequired = $false
        })
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Voice configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update Voice configuration: $_" -Tag "Error"
    }
}

function Set-CertificateConfiguration {
    $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/x509Certificate"

    $body = @{
        "@odata.type" = "#microsoft.graph.x509CertificateAuthenticationMethodConfiguration"
        state = "disabled"
        includeTargets = @(@{
            targetType = "group"
            id = "all_users"
            isRegistrationRequired = $false
        })
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Certificate configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update Certificate configuration: $_" -Tag "Error"
    }
}

# ---------------------------[ User Settings ]---------------------------

function Set-UserConfiguration {
    $uri = "https://graph.microsoft.com/v1.0/policies/authorizationPolicy"

    $body = @{
        defaultUserRolePermissions = @{
            allowedToCreateApps = $false
            allowedToCreateSecurityGroups = $false
            allowedToCreateTenants = $false
            allowedToReadBitlockerKeysForOwnedDevice = $false
        }
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "User configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update User configuration: $_" -Tag "Error"
    }
}

# ---------------------------[ Guest Settings ]---------------------------

function Set-GuestConfiguration {
    $uri = "https://graph.microsoft.com/v1.0/policies/authorizationPolicy"

    $body = @{
        allowInvitesFrom = "adminsGuestInvitersAndAllMembers"
        guestUserRoleId = "2af84b1e-32c8-42b7-82bc-daa82404023b"
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Guest configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update Guest configuration: $_" -Tag "Error"
    }
}

# ---------------------------[ Group Settings ]---------------------------

function New-Group {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Name
    )
    $checkUri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$Name' and securityEnabled eq true"
    $createUri = "https://graph.microsoft.com/v1.0/groups"

    try {
        $existingGroup = Invoke-MgGraphRequest -Method GET -Uri $checkUri -ErrorAction Stop

        if ($existingGroup.value.Count -gt 0) {
            $groupId = $existingGroup.value[0].id
            # Write-Log "Group '$Name' already exists. ID: $groupId" -Tag "Info"
            return $groupId
        }

        $body = @{
            displayName     = $Name
            mailEnabled     = $false
            mailNickname    = $Name.Replace(" ", "").ToLower()
            securityEnabled = $true
            groupTypes      = @()
        }

        Write-Log "Creating group '$Name'..." -Tag "Info"
        $createdGroup = Invoke-MgGraphRequest -Method POST -Uri $createUri -Body ($body | ConvertTo-Json -Depth 10) -ContentType "application/json" -ErrorAction Stop
        $groupId = $createdGroup.id
        Write-Log "Group '$Name' created successfully. ID: $groupId" -Tag "Success"
        return $groupId
    } catch {
        Write-Log "Error: $_" -Tag "Error"
        return $null
    }
}

function Get-GroupUnifiedId {
    $uri = "https://graph.microsoft.com/v1.0/groupSettings"

    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $groupUnifiedSetting = $response.value | Where-Object { $_.displayName -eq "Group.Unified" }

        if (-not $groupUnifiedSetting) {
            Write-Log "Group.Unified setting not found." -Tag "Error"
            return $null
        }

        $settingId = $groupUnifiedSetting.id
        # Write-Log "Found Group.Unified setting ID: $settingId" -Tag "Debug"
        return $settingId
    } catch {
        Write-Log "Failed to retrieve group settings: $_" -Tag "Error"
        return $null
    }
}

function Set-GroupConfiguration {
    $uri = "https://graph.microsoft.com/v1.0/groupSettings/$groupUnifiedId"

    $body = @{
        values = @(
            @{ name = "EnableGroupCreation"; value = "false" },
            @{ name = "GroupCreationAllowedGroupId"; value = $groupCreationGroupId },
            @{ name = "EnableMIPLabels"; value = "true" }
        )
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Group configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update Group configuration: $_" -Tag "Error"
    }
}

# ---------------------------[ Application Consent ]---------------------------

function Set-ApplicationUserConsentConfiguration {
    $uri = "https://graph.microsoft.com/v1.0/policies/authorizationPolicy"

    try {
        $currentPolicy = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    } catch {
        Write-Log "Failed to retrieve authorizationPolicy: $_" -Tag "Error"
        return
    }

    $currentPermissions = $currentPolicy.defaultUserRolePermissions.permissionGrantPoliciesAssigned

    if (-not $currentPermissions) {
        $currentPermissions = @()
    }

    $expected = "ManagePermissionGrantsForSelf.microsoft-user-default-low"
    $legacy   = "ManagePermissionGrantsForSelf.microsoft-user-default-legacy"

    # Remove legacy if present
    $cleaned = @($currentPermissions | Where-Object { $_ -ne $legacy })

    # Skip update if expected is already present and no change is needed
    if ($cleaned -contains $expected -and $cleaned.Count -eq $currentPermissions.Count) {
        Write-Log "Required Application User Consent configuration already present." -Tag "Info"
        return
    }

    # Add expected if not present
    if ($cleaned -notcontains $expected) {
        $cleaned += $expected
    }

    $body = @{
        defaultUserRolePermissions = @{
            permissionGrantPoliciesAssigned = $cleaned
        }
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Application User Consent configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update Application User Consent configuration: $_" -Tag "Error"
    }
}

function Set-ApplicationAdminConsentConfiguration {
    $uri = "https://graph.microsoft.com/v1.0/policies/adminConsentRequestPolicy"

    $body = @{
        isEnabled = $true
        notifyReviewers = $true
        remindersEnabled = $true
        requestDurationInDays = "30"
        reviewers = @(
            @{  
                query = "/beta/roleManagement/directory/roleAssignments?`$filter=roleDefinitionId eq '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3'"
                queryType = "MicrosoftGraph"
            },
            @{
                query = "/beta/roleManagement/directory/roleAssignments?`$filter=roleDefinitionId eq '62e90394-69f5-4237-9190-012177145e10'"
                queryType = "MicrosoftGraph"
            }
        )
    }

    $json = $body | ConvertTo-Json -Depth 10 -Compress

    try {
        Invoke-MgGraphRequest -Method PUT -Uri $uri -Body $json -ContentType "application/json" -ErrorAction Stop
        Write-Log "Application Consent configuration updated successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to update Application Consent configuration: $_" -Tag "Error"
    }
}

# ---------------------------[ Execution ]---------------------------

Set-DeviceSettings
Set-Fido2Configuration
Set-MicrosoftAuthenticatorConfiguration
Set-TemporaryAccessPassConfiguration
Set-SoftwareOathConfiguration
Set-HardwareOathConfiguration
Set-EmailConfiguration
Set-SmsConfiguration
Set-VoiceConfiguration
Set-CertificateConfiguration
Set-UserConfiguration
Set-GuestConfiguration
$groupCreationGroupId = New-Group -Name "Group Creation - Users"
$groupUnifiedId = Get-GroupUnifiedId
Set-GroupConfiguration
Set-ApplicationUserConsentConfiguration
Set-ApplicationAdminConsentConfiguration

# ---------------------------[ Script Complete ]---------------------------

Complete-Script -ExitCode 0