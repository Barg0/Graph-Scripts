# ---------------------------[ Script Start Timestamp ]---------------------------

# Capture start time to log script duration
$scriptStartTime = Get-Date

# ---------------------------[ Configuration ]---------------------------

# Mailbox UPN where personal contacts will be created/updated
$targetMailboxUpn = "contacts@yourdomain.com"

# App-only Graph authentication (client secret)
# REQUIRED APP PERMISSIONS (Application):
#   - OrgContact.Read.All   (read org contacts /contacts)
#   - Contacts.ReadWrite    (manage personal contacts in user mailboxes)
$tenantId     = "<YOUR-TENANT-ID>"
$clientId     = "<YOUR-APP-CLIENT-ID>"
$clientSecret = "<YOUR-CLIENT-SECRET>"   # Use a secure store in production

# ---------------------------[ Script name ]---------------------------

# Script name used for folder/log naming
$scriptName  = "Sync-MailContacts"
$logFileName = "$($scriptName).log"

# ---------------------------[ Logging Setup ]---------------------------

# Logging control switches
$log                   = $true      # Set to $false to disable logging in shell
$enableLogFile         = $false     # Set to $false to disable file output
$enableComparisonDebug = $false     # Set to $true to log property-level comparisons

# Define the log output location
$logFileDirectory = "$PSScriptRoot"
$logFile          = "$logFileDirectory\$logFileName"

# Ensure the log directory exists
if ($enableLogFile -and -not (Test-Path $logFileDirectory)) {
    New-Item -ItemType Directory -Path $logFileDirectory -Force | Out-Null
}

# Function to write structured logs to file and console
function Write-Log {
    param (
        [string]$Message,
        [string]$Tag = "Info"
    )

    if (-not $log) {
        return
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $tagList   = @("Start", "Check", "Info", "Success", "Error", "Debug", "End")
    $rawTag    = $Tag.Trim()

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

# ---------------------------[ Exit Function ]---------------------------

function Complete-Script {
    param(
        [int]$ExitCode
    )

    $scriptEndTime = Get-Date
    $duration      = $scriptEndTime - $scriptStartTime

    Write-Log "Script execution time: $($duration.ToString("hh\:mm\:ss\.ff"))" -Tag "Info"
    Write-Log "Exit Code: $ExitCode" -Tag "Info"
    Write-Log "======== Script Completed ========" -Tag "End"
    exit $ExitCode
}

# ---------------------------[ Start ]---------------------------

Write-Log "======== Script Started ========" -Tag "Start"
Write-Log "ComputerName: $env:COMPUTERNAME | User: $env:USERNAME | Script: $scriptName" -Tag "Info"

if ([string]::IsNullOrWhiteSpace($targetMailboxUpn)) {
    Write-Log "Target mailbox UPN is not configured. Please set `$targetMailboxUpn in the configuration section." -Tag "Error"
    Complete-Script -ExitCode 1
}

if ([string]::IsNullOrWhiteSpace($tenantId) -or
    [string]::IsNullOrWhiteSpace($clientId) -or
    [string]::IsNullOrWhiteSpace($clientSecret)) {

    Write-Log "TenantId / ClientId / ClientSecret are not fully configured. Please set them in the configuration section." -Tag "Error"
    Complete-Script -ExitCode 1
}

# ---------------------------[ Connect to Microsoft Graph (App Only) ]---------------------------

function Test-MicrosoftGraphConnection {

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

    try {
        Write-Log "Connecting to Microsoft Graph using app-only (client secret)..." -Tag "Info"

        $clientSecretSecure     = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
        $clientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $clientId, $clientSecretSecure

        Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $clientSecretCredential | Out-Null

        $context = Get-MgContext
        Write-Log "Connected to Microsoft Graph as app with ClientId $($context.ClientId) in Tenant $($context.TenantId)." -Tag "Success"
    } catch {
        Write-Log "Failed to connect to Microsoft Graph (app-only): $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

# ---------------------------[ Helper Functions ]---------------------------

function Get-DirectoryOrgContacts {
    <#
        .SYNOPSIS
        Retrieves all organizational contacts (orgContact) from Microsoft Graph.
    #>
    try {
        Write-Log "Retrieving organizational contacts (orgContact) from Microsoft Graph..." -Tag "Debug"

        $orgContacts = Get-MgContact -All -Property `
            "id,displayName,givenName,surname,companyName,department,jobTitle,mail,mailNickname,addresses,phones,imAddresses"

        Write-Log "Retrieved $($orgContacts.Count) orgContacts from directory." -Tag "Success"
        return $orgContacts
    } catch {
        Write-Log "Failed to retrieve orgContacts: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

function Get-MailboxContactIndex {
    <#
        .SYNOPSIS
        Retrieves all personal contacts from the target mailbox via Graph
        and returns a hashtable indexed by primary email address (lowercase).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        Write-Log "Retrieving existing contacts from mailbox '$UserId' via Microsoft Graph..." -Tag "Debug"
        $mailboxContacts = Get-MgUserContact -UserId $UserId -All

        $contactIndex = @{}

        foreach ($contact in $mailboxContacts) {
            if ($null -eq $contact.EmailAddresses -or $contact.EmailAddresses.Count -eq 0) {
                continue
            }

            $primaryEmail = $contact.EmailAddresses[0].Address
            if ([string]::IsNullOrWhiteSpace($primaryEmail)) {
                continue
            }

            $key = $primaryEmail.ToLowerInvariant()
            if (-not $contactIndex.ContainsKey($key)) {
                $contactIndex[$key] = $contact
            }
        }

        Write-Log "Retrieved $($mailboxContacts.Count) existing contacts from mailbox and indexed $($contactIndex.Count) by email." -Tag "Success"

        return $contactIndex
    } catch {
        Write-Log "Failed to retrieve existing mailbox contacts: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

function Get-DesiredContactBody {
    <#
        .SYNOPSIS
        Builds the desired Graph contact body from an orgContact.
        Only non-empty properties are included, to avoid wiping fields unnecessarily.
    #>
    param(
        [Parameter(Mandatory = $true)]
        $OrgContact
    )

    $primaryEmail = $OrgContact.Mail
    $contactBody  = @{}

    # Basic identity
    if (-not [string]::IsNullOrWhiteSpace($OrgContact.DisplayName)) {
        $contactBody["displayName"] = $OrgContact.DisplayName
    }

    if (-not [string]::IsNullOrWhiteSpace($OrgContact.GivenName)) {
        $contactBody["givenName"] = $OrgContact.GivenName
    }

    if (-not [string]::IsNullOrWhiteSpace($OrgContact.Surname)) {
        $contactBody["surname"]  = $OrgContact.Surname
    }

    # nickname from MailNickname
    if (-not [string]::IsNullOrWhiteSpace($OrgContact.MailNickname)) {
        $contactBody["nickName"] = $OrgContact.MailNickname
    }

    # Organization
    if (-not [string]::IsNullOrWhiteSpace($OrgContact.CompanyName)) {
        $contactBody["companyName"] = $OrgContact.CompanyName
    }

    if (-not [string]::IsNullOrWhiteSpace($OrgContact.Department)) {
        $contactBody["department"] = $OrgContact.Department
    }

    if (-not [string]::IsNullOrWhiteSpace($OrgContact.JobTitle)) {
        $contactBody["jobTitle"] = $OrgContact.JobTitle
    }

    # Phones from orgContact.phones (type: mobile, business, etc.)
    $businessPhones = @()
    $mobilePhone    = $null

    if ($null -ne $OrgContact.Phones -and $OrgContact.Phones.Count -gt 0) {
        foreach ($phone in $OrgContact.Phones) {
            $number = $phone.Number
            $type   = $phone.Type

            if ([string]::IsNullOrWhiteSpace($number)) {
                continue
            }

            switch ($type) {
                "business" {
                    $businessPhones += $number
                }
                "mobile" {
                    if ($null -eq $mobilePhone) {
                        $mobilePhone = $number
                    }
                }
                default { }
            }
        }
    }

    if ($businessPhones.Count -gt 0) {
        $contactBody["businessPhones"] = $businessPhones
    }

    if (-not [string]::IsNullOrWhiteSpace($mobilePhone)) {
        $contactBody["mobilePhone"] = $mobilePhone
    }

    # Email
    if (-not [string]::IsNullOrWhiteSpace($primaryEmail)) {
        $contactBody["emailAddresses"] = @(
            @{
                name    = $OrgContact.DisplayName
                address = $primaryEmail
            }
        )
    }

    # Business address from orgContact.addresses[0]
    $businessAddress = $null
    if ($null -ne $OrgContact.Addresses -and $OrgContact.Addresses.Count -gt 0) {
        $businessAddress = $OrgContact.Addresses[0]
    }

    if ($null -ne $businessAddress) {
        $street        = $businessAddress.Street
        $city          = $businessAddress.City
        $state         = $businessAddress.State
        $postalCode    = $businessAddress.PostalCode
        $countryRegion = $businessAddress.CountryOrRegion

        if (
            -not [string]::IsNullOrWhiteSpace($street)      -or
            -not [string]::IsNullOrWhiteSpace($city)        -or
            -not [string]::IsNullOrWhiteSpace($state)       -or
            -not [string]::IsNullOrWhiteSpace($postalCode)  -or
            -not [string]::IsNullOrWhiteSpace($countryRegion)
        ) {
            $contactBody["businessAddress"] = @{
                street          = $street
                city            = $city
                state           = $state
                postalCode      = $postalCode
                countryOrRegion = $countryRegion
            }
        }
    }

    # imAddresses
    $imList = @()

    if ($null -ne $OrgContact.ImAddresses -and $OrgContact.ImAddresses.Count -gt 0) {
        foreach ($im in $OrgContact.ImAddresses) {
            if (-not [string]::IsNullOrWhiteSpace($im)) {
                $imList += $im
            }
        }
    } elseif ($null -ne $OrgContact.AdditionalProperties -and
              $OrgContact.AdditionalProperties.ContainsKey("imAddresses")) {

        $imValues = $OrgContact.AdditionalProperties["imAddresses"]
        foreach ($im in $imValues) {
            if (-not [string]::IsNullOrWhiteSpace($im)) {
                $imList += $im
            }
        }
    }

    if ($imList.Count -gt 0) {
        $contactBody["imAddresses"] = $imList
    }

    return $contactBody
}

function Get-ContactUpdateBody {
    <#
        .SYNOPSIS
        Compares desired contact body with an existing Graph contact and
        returns a hashtable containing only the properties that need to be updated.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$DesiredBody,

        [Parameter(Mandatory = $true)]
        $ExistingContact
    )

    $updateBody = @{}

    function Convert-ToNormalizedString {
        param([string]$Value)
        if ([string]::IsNullOrWhiteSpace($Value)) {
            return ""
        }
        return $Value.Trim()
    }

    # Explicit map for scalar properties -> Graph contact properties
    $scalarPropertyMap = @{
        displayName   = "DisplayName"
        givenName     = "GivenName"
        surname       = "Surname"
        companyName   = "CompanyName"
        department    = "Department"
        jobTitle      = "JobTitle"
        mobilePhone   = "MobilePhone"
        nickName      = "NickName"
    }

    foreach ($property in $DesiredBody.Keys) {

        switch ($property) {

            # --- Arrays (businessPhones) ---
            "businessPhones" {
                $desiredString = -join ($DesiredBody[$property] | ForEach-Object {
                    Convert-ToNormalizedString -Value $_
                })

                $existingArray  = $ExistingContact.BusinessPhones
                $existingString = -join ($existingArray | ForEach-Object {
                    Convert-ToNormalizedString -Value $_
                })

                if ($enableComparisonDebug) {
                    Write-Log "Compare businessPhones: Desired='$desiredString' Existing='$existingString'" -Tag "Debug"
                }

                if ($desiredString -ne $existingString) {
                    $updateBody[$property] = $DesiredBody[$property]
                }
            }

            # --- Arrays (imAddresses) ---
            "imAddresses" {
                $desiredString = -join ($DesiredBody[$property] | ForEach-Object {
                    Convert-ToNormalizedString -Value $_
                })

                $existingArray  = $ExistingContact.ImAddresses
                $existingString = -join ($existingArray | ForEach-Object {
                    Convert-ToNormalizedString -Value $_
                })

                if ($enableComparisonDebug) {
                    Write-Log "Compare imAddresses: Desired='$desiredString' Existing='$existingString'" -Tag "Debug"
                }

                if ($desiredString -ne $existingString) {
                    $updateBody[$property] = $DesiredBody[$property]
                }
            }

            # --- emailAddresses (compare primary) ---
            "emailAddresses" {
                $desiredEmail = $null
                if ($DesiredBody[$property].Count -gt 0) {
                    $desiredEmail = Convert-ToNormalizedString -Value $DesiredBody[$property][0].address
                }

                $existingEmail = $null
                if ($ExistingContact.EmailAddresses.Count -gt 0) {
                    $existingEmail = Convert-ToNormalizedString -Value $ExistingContact.EmailAddresses[0].Address
                }

                if ($enableComparisonDebug) {
                    Write-Log "Compare emailAddresses: Desired='$desiredEmail' Existing='$existingEmail'" -Tag "Debug"
                }

                if ($desiredEmail -ne $existingEmail) {
                    $updateBody[$property] = $DesiredBody[$property]
                }
            }

            # --- businessAddress (complex) ---
            "businessAddress" {
                $desiredAddress  = $DesiredBody[$property]
                $existingAddress = $ExistingContact.BusinessAddress

                $desiredString = @(
                    Convert-ToNormalizedString -Value $desiredAddress.street
                    Convert-ToNormalizedString -Value $desiredAddress.city
                    Convert-ToNormalizedString -Value $desiredAddress.state
                    Convert-ToNormalizedString -Value $desiredAddress.postalCode
                    Convert-ToNormalizedString -Value $desiredAddress.countryOrRegion
                ) -join "|"

                $existingString = ""
                if ($null -ne $existingAddress) {
                    $existingString = @(
                        Convert-ToNormalizedString -Value $existingAddress.Street
                        Convert-ToNormalizedString -Value $existingAddress.City
                        Convert-ToNormalizedString -Value $existingAddress.State
                        Convert-ToNormalizedString -Value $existingAddress.PostalCode
                        Convert-ToNormalizedString -Value $existingAddress.CountryOrRegion
                    ) -join "|"
                }

                if ($enableComparisonDebug) {
                    Write-Log "Compare businessAddress: Desired='$desiredString' Existing='$existingString'" -Tag "Debug"
                }

                if ($desiredString -ne $existingString) {
                    $updateBody[$property] = $DesiredBody[$property]
                }
            }

            # --- Scalars: use explicit property map ---
            default {
                $graphPropertyName = $scalarPropertyMap[$property]
                if ([string]::IsNullOrWhiteSpace($graphPropertyName)) {
                    $graphPropertyName = $property.Substring(0,1).ToUpper() + $property.Substring(1)
                }

                $desiredValue  = Convert-ToNormalizedString -Value $DesiredBody[$property]
                $existingValue = Convert-ToNormalizedString -Value $ExistingContact.$graphPropertyName

                if ($enableComparisonDebug) {
                    Write-Log "Compare $property ($graphPropertyName): Desired='$desiredValue' Existing='$existingValue'" -Tag "Debug"
                }

                if ($desiredValue -ne $existingValue) {
                    $updateBody[$property] = $DesiredBody[$property]
                }
            }
        }
    }

    return $updateBody
}

function Invoke-OrgContactSync {
    <#
        .SYNOPSIS
        Main sync routine: reads orgContacts from Graph and syncs them into a mailbox as personal contacts.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetMailboxUpn
    )

    Write-Log "Starting sync of orgContacts into mailbox '$TargetMailboxUpn'..." -Tag "Info"

    $orgContacts         = Get-DirectoryOrgContacts
    $mailboxContactIndex = Get-MailboxContactIndex -UserId $TargetMailboxUpn

    $totalContacts = $orgContacts.Count
    $createdCount  = 0
    $updatedCount  = 0
    $skippedCount  = 0
    $errorCount    = 0

    foreach ($orgContact in $orgContacts) {

        $primaryEmail = $orgContact.Mail
        if ([string]::IsNullOrWhiteSpace($primaryEmail)) {
            Write-Log "orgContact '$($orgContact.Id)' has no Mail property. Skipping." -Tag "Debug"
            $skippedCount++
            continue
        }

        $emailKey    = $primaryEmail.ToLowerInvariant()
        $desiredBody = Get-DesiredContactBody -OrgContact $orgContact

        if (-not $mailboxContactIndex.ContainsKey($emailKey)) {

            try {
                Write-Log "Creating new contact for '$primaryEmail' in mailbox '$TargetMailboxUpn'." -Tag "Info"
                New-MgUserContact -UserId $TargetMailboxUpn -BodyParameter $desiredBody | Out-Null
                $createdCount++
            } catch {
                Write-Log "Failed to create contact for '$primaryEmail': $_" -Tag "Error"
                $errorCount++
            }

        } else {

            $existingContact = $mailboxContactIndex[$emailKey]
            $updateBody      = Get-ContactUpdateBody -DesiredBody $desiredBody -ExistingContact $existingContact

            if ($updateBody.Keys.Count -eq 0) {
                Write-Log "Contact '$primaryEmail' already up to date. Skipping." -Tag "Debug"
                $skippedCount++
                continue
            }

            try {
                Write-Log "Updating contact '$primaryEmail' in mailbox '$TargetMailboxUpn' (fields: $($updateBody.Keys -join ', '))." -Tag "Info"
                Update-MgUserContact -UserId $TargetMailboxUpn -ContactId $existingContact.Id -BodyParameter $updateBody | Out-Null
                $updatedCount++
            } catch {
                Write-Log "Failed to update contact '$primaryEmail': $_" -Tag "Error"
                $errorCount++
            }
        }
    }

    Write-Log "Sync summary: Total=$totalContacts | Created=$createdCount | Updated=$updatedCount | Skipped=$skippedCount | Errors=$errorCount" -Tag "Success"
}

# ---------------------------[ Main Execution ]---------------------------

Test-MicrosoftGraphConnection

Invoke-OrgContactSync -TargetMailboxUpn $targetMailboxUpn

Complete-Script -ExitCode 0