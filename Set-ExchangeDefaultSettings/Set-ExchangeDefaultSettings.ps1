# Script version:   2025-06-01 12:40
# Script author:    Barg0

# ---------------------------[ Script Start Timestamp ]---------------------------

# Capture start time to log script duration
$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$graphScopes = @(
    "Domain.Read.All"
)

# ---------------------------[ Script name ]---------------------------

# Script name used for folder/log naming
$scriptName = "Set-ExchangeDefaultSettings"
$logFileName = "$($scriptName).log"

# ---------------------------[ Logging Setup ]---------------------------

# Logging control switches
$log = $true                     # Set to $false to disable logging in shell
$enableLogFile = $false          # Set to $false to disable file output

# Define the log output location
$logFileDirectory = "$PSScriptRoot"
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

# ---------------------------[ Start ]---------------------------

Write-Log "======== Script Started ========" -Tag "Start"
Write-Log "ComputerName: $env:COMPUTERNAME | User: $env:USERNAME | Script: $scriptName" -Tag "Info"


# ---------------------------[ Connect to Exchange Online ]---------------------------

function Test-ExchangeOnlineConnection {
    # Ensure Exchange Online PowerShell Module is present
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log "Installing Exchange Online PowerShell module..." -Tag "Info"
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
    }

    try {
        Import-Module ExchangeOnlineManagement -Force
        if (-not (Get-ConnectionInformation | Where-Object { $_.Name -match 'ExchangeOnline' -and $_.State -eq 'Connected' })) {
            Write-Log "Connecting to Exchange Online" -Tag "Info"
            Connect-ExchangeOnline *>&1 | Out-Null
            Write-Log "Connected to Exchange Online." -Tag "Success"
        } else {
            Write-Log "Already connected to Exchange Online." -Tag "Info"
        }
    } catch {
        Write-Log "Failed to connect to Exchange Online. $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

# ---------------------------[ Connect to Microsoft Graph ]---------------------------

function Test-MicrosoftGraphConnection {
    # Test Connection
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

    # Connect to Graph
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
}

# ---------------------------[ Get Domains ]---------------------------

function Get-DefaultDomain {
    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/domains"

        $defaultDomain = $response.value | Where-Object { $_.isDefault -eq $true }

        if ($null -eq $defaultDomain) {
            Write-Log "No default domain found." -Tag "Error"
            return $null
        }

        Write-Log "Default domain is: $($defaultDomain.id)" -Tag "Success"
        return $defaultDomain.id
    }
    catch {
        Write-Log "Failed to retrieve default domain: $_" -Tag "Error"
        return $null
    }
}

function Get-AllDomains {
    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/domains"

        $allDomains = $response.value

        if ($null -eq $allDomains) {
            Write-Log "No Domains found." -Tag "Error"
            return $null
        }

        Write-Log "Domains: $($allDomains.id)" -Tag "Success"
        return $allDomains.id
    }
    catch {
        Write-Log "Failed to retrieve default domain: $_" -Tag "Error"
        return $null
    }
}

# ---------------------------[ Create Shared Mailbox ]---------------------------

function New-SharedMailbox {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        [Parameter(Mandatory = $true)]
        [string]$MailAlias,
        [Parameter(Mandatory = $true)]
        [string]$Language,
        [Parameter(Mandatory = $true)]
        [bool]$VisibleInGal
    )

    try {
        $userPrincipalName = "$($MailAlias)@$($defaultDomain)"
        $existingMailbox = Get-Mailbox -Identity $userPrincipalName -ErrorAction SilentlyContinue
        if ($null -ne $existingMailbox) {
            Write-Log "Shared mailbox '$userPrincipalName' already exists." -Tag "Info"
            return $userPrincipalName
        }

        # Write-Log "Creating shared mailbox '$userPrincipalName'..." -Tag "Info"
        New-Mailbox -Shared -Name $MailAlias -DisplayName $DisplayName -Alias $MailAlias -ErrorAction Stop | Out-Null

        $checkMailbox = Get-Mailbox -Identity $userPrincipalName -ErrorAction SilentlyContinue
        if ($null -ne $checkMailbox) {
            Write-Log "Shared mailbox '$userPrincipalName' created successfully." -Tag "Success"
        } else {
            Write-Log "Shared mailbox '$userPrincipalName' was not found after creation attempt." -Tag "Error"
            return $false
        }

        # Hide from GAL
        try {
            if (!$VisibleInGal){
            Set-Mailbox -Identity $userPrincipalName -HiddenFromAddressListsEnabled $true -ErrorAction Stop | Out-Null
            }
        } catch {
            Write-Log "Failed to set GAL visibility for $primarySmtp. $_" -Tag "Error"
        }        

        # Set Language
        try {
            Set-MailboxRegionalConfiguration -Identity $userPrincipalName -Language $Language -TimeZone "W. Europe Standard Time" -DateFormat "dd.MM.yyyy" -TimeFormat "HH:mm" -LocalizeDefaultFolderName -ErrorAction Stop | Out-Null
            # Write-Log "Set $userPrincipalName to $Language" -Tag "Info"
        } catch {
            Write-Log "Failed to set regional settings on $userPrincipalName. $_" -Tag "Error"
        }
        return $userPrincipalName
    }
    catch {
        Write-Log "Failed to create shared mailbox '$userPrincipalName': $_" -Tag "Error"
        return $false
    }
}

# ---------------------------[ Organization Customization ]---------------------------

function Test-OrganziationCustomization {
    try {
        $orgConfig = Get-OrganizationConfig
        if ($orgConfig.IsDehydrated) {
            Write-Log "Organization is dehydrated. Enabling customization..." -Tag "Info"
            Enable-OrganizationCustomization -ErrorAction Stop
            Start-Sleep -Seconds 60

            $orgConfig = Get-OrganizationConfig
            if ($orgConfig.IsDehydrated) {
                Write-Log "Customization failed to apply. Still dehydrated." -Tag "Error"
                Complete-Script -ExitCode 1
            }

            Write-Log "Customization successfully enabled." -Tag "Success"
        } else {
            Write-Log "Organization is already customized." -Tag "Success"
        }

        return $true
    }
    catch {
        Write-Log "Exception encountered while enabling customization: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

# ---------------------------[ Mailbox Import Export Role ]---------------------------

function Test-MailboxImportExportRole {

    $roleGroup = "Organization Management"
    $roleName = "Mailbox Import Export"

    # Write-Log "Checking if '$roleName' is assigned to '$roleGroup'..." -Tag "Debug"

    try {
        $assignedRoles = Get-ManagementRoleAssignment -RoleAssignee $roleGroup -ErrorAction Stop | Select-Object -ExpandProperty Role

        if ($assignedRoles -contains $roleName) {
            Write-Log "Role '$roleName' is already assigned to '$roleGroup'." -Tag "Info"
            return $true
        }

        Write-Log "Assigning role '$roleName' to '$roleGroup'..." -Tag "Debug"

        New-ManagementRoleAssignment -Role $roleName -SecurityGroup $roleGroup -ErrorAction Stop | Out-Null

        Start-Sleep -Seconds 10

        # Re-check after assignment
        $assignedRoles = Get-ManagementRoleAssignment -RoleAssignee $roleGroup -ErrorAction Stop | Select-Object -ExpandProperty Role

        if ($assignedRoles -contains $roleName) {
            Write-Log "Role '$roleName' successfully assigned to '$roleGroup'." -Tag "Success"
            return $true
        } else {
            Write-Log "Role '$roleName' was not found after assignment attempt." -Tag "Error"
            return $false
        }
    }
    catch {
        Write-Log "Error while checking or assigning role '$roleName': $_" -Tag "Error"
        return $false
    }
}

# ---------------------------[ Quarantine Policy ]---------------------------

function New-LimitedAccessQuarantinePolicy {
    $policyName = "LimitedAccessPolicy"
    # Write-Log "Checking for existing policy '$policyName'..." -Tag "Debug"

    try {
        $existing = Get-QuarantinePolicy -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $policyName }
        if ($existing) {
            Write-Log "Policy '$policyName' already exists." -Tag "Info"
            return $true
        }
        New-QuarantinePolicy -Name $policyName -EndUserQuarantinePermissionsValue 155 | Out-Null
        Set-QuarantinePolicy -Identity $policyName -ESNEnabled $true -IncludeMessagesFromBlockedSenderAddress $true | Out-Null

        Write-Log "Policy '$policyName' created successfully." -Tag "Success"
        return $true
    }
    catch {
        Write-Log "Error creating '$policyName': $_" -Tag "Error"
        return $false
    }
}

function New-FullAccessQuarantinePolicy {
    $policyName = "FullAccessPolicy"
    # Write-Log "Checking for existing policy '$policyName'..." -Tag "Debug"

    try {
        $existing = Get-QuarantinePolicy -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $policyName }
        if ($existing) {
            Write-Log "Policy '$policyName' already exists." -Tag "Info"
            return $true
        }

        New-QuarantinePolicy -Name $policyName -EndUserQuarantinePermissionsValue 183 | Out-Null
        Set-QuarantinePolicy -Identity $policyName -ESNEnabled $true -IncludeMessagesFromBlockedSenderAddress $true | Out-Null

        Write-Log "Policy '$policyName' created successfully." -Tag "Success"
        return $true
    }
    catch {
        Write-Log "Error creating '$policyName': $_" -Tag "Error"
        return $false
    }
}

# ---------------------------[ Global Notification Policy ]---------------------------

function Set-GlobalQuarantineNotificationPolicy {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SenderAddress,
        [Parameter(Mandatory = $true)]
        [string]$Language
    )

    try {
        $globalPolicy = Get-QuarantinePolicy -QuarantinePolicyType GlobalQuarantinePolicy -ErrorAction Stop

        Set-QuarantinePolicy -Identity $globalPolicy.Identity `
            -MultiLanguageCustomDisclaimer @("") `
            -MultiLanguageSenderName @("") `
            -EsnCustomSubject @("") `
            -MultiLanguageSetting @($Language) `
            -EndUserSpamNotificationCustomFromAddress $SenderAddress `
            -OrganizationBrandingEnabled $true `
            -EndUserSpamNotificationFrequency "04:00:00" `
            -ErrorAction Stop

        Write-Log "Global quarantine notification policy updated successfully." -Tag "Success"
        return $true
    }
    catch {
        Write-Log "Failed to update global quarantine notification policy: $_" -Tag "Error"
        return $false
    }
}

# ---------------------------[ Execution ]---------------------------

Test-ExchangeOnlineConnection
Test-MicrosoftGraphConnection
$defaultDomain = Get-DefaultDomain
Get-AllDomains
Test-OrganziationCustomization | Out-Null
Test-MailboxImportExportRole | Out-Null
New-LimitedAccessQuarantinePolicy | Out-Null
New-FullAccessQuarantinePolicy | Out-Null
$sharedMailboxMicrosoftDefender = New-SharedMailbox -DisplayName "Microsoft Defender" -MailAlias "microsoft-defender" -Language "de-DE" -VisibleInGal $false
Set-GlobalQuarantineNotificationPolicy -SenderAddress $sharedMailboxMicrosoftDefender -Language "German" | Out-Null

# ---------------------------[ End ]---------------------------

Complete-Script -ExitCode 0