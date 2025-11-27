# ---------------------------[ Script Start Timestamp ]---------------------------

# Capture start time to log script duration
$scriptStartTime = Get-Date

# ---------------------------[ Configuration ]---------------------------

# Mailbox UPN where personal contacts will be created/updated
$targetMailboxUpn = "contacts@yourdomain.com"

# App-only Graph authentication (client secret)
# REQUIRED APP PERMISSIONS (Application):
#   - Directory.Read.All    (read org contacts /contacts)
#   - Contacts.ReadWrite    (manage personal contacts in user mailboxes)
$tenantId     = "<YOUR-TENANT-ID>"
$clientId     = "<YOUR-APP-CLIENT-ID>"
$clientSecret = "<YOUR-CLIENT-SECRET>"   # Use a secure store in production

# Enable deletion of mailbox contacts that no longer exist as orgContacts
$enableDeletion = $true   # Set to $false if you want to disable deletions

# Email reporting configuration
$enableEmailReport    = $true                     # Set to $false to disable email reporting
$reportSmtpServer     = "smtp.yourdomain.com"     # SMTP server
$reportSmtpPort       = 587                       # SMTP port (25, 587, 2525, etc.)
$reportUseSsl         = $true                     # Use SSL/TLS for SMTP
$reportFrom           = "noreply@yourdomain.com"  # Sender address
$reportTo             = "admin@yourdomain.com"    # Recipient(s), comma-separated if multiple
$reportSubject        = "Contact sync report"
$reportSmtpCredential = $null                     # Optional: set to Get-Credential if auth is required

# ---------------------------[ Script name ]---------------------------

# Script name used for folder/log naming
$scriptName  = "Sync-MailContacts"
$logFileName = "$scriptName.log"

# ---------------------------[ Logging Setup ]---------------------------

# Logging control switches
$log        = $true   # Set to $false to disable ALL console logging
$enableLogFile = $false   # Set to $false to disable file logging
$logDebug   = $false  # Set to $true to enable debug logging

# Define the log output location
$logFileDirectory = "$PSScriptRoot"
$logFile          = "$logFileDirectory\$logFileName"

# Ensure the log directory exists
if ($enableLogFile -and -not (Test-Path $logFileDirectory)) {
    New-Item -ItemType Directory -Path $logFileDirectory -Force | Out-Null
}

# Function to write structured logs to file and console
function Write-Log {
    [CmdletBinding()]
    param (
        [string]$Message,
        [string]$Tag = "Info"
    )

    if (-not $log) { return }

    # Suppress debug entries when $logDebug = $false
    if ($Tag -eq "Debug" -and -not $logDebug) { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $tagList   = @("Start", "Check", "Info", "Success", "Error", "Debug", "End")

    $rawTag = if ($tagList -contains $Tag) { $Tag.PadRight(7) } else { "Error  " }

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

    # Write to file
    if ($enableLogFile) {
        $logMessage | Out-File -FilePath $logFile -Append
    }

    # Write to console
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
    $duration      = $scriptEndTime - $scriptStartTime

    Write-Log "Script execution time: $($duration.ToString('hh\:mm\:ss\.ff'))" -Tag "Info"
    Write-Log "Exit Code: $ExitCode" -Tag "Info"
    Write-Log "======== Script Completed ========" -Tag "End"

    exit $ExitCode
}

# ---------------------------[ Script Start ]---------------------------

Write-Log "======== Script Started ========" -Tag "Start"
Write-Log "ComputerName: $env:COMPUTERNAME | User: $env:USERNAME | Script: $scriptName" -Tag "Info"

# Basic config validation
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

# ---------------------------[ Helper: Email Report ]---------------------------

function Send-ContactSyncReport {
    param(
        [int]$TotalContacts,
        [int]$CreatedCount,
        [int]$UpdatedCount,
        [int]$SkippedCount,
        [int]$DeletedCount,
        [int]$ErrorCount,
        [string[]]$CreatedContacts,
        [string[]]$UpdatedContacts,
        [string[]]$SkippedContacts,
        [string[]]$DeletedContacts,
        [string[]]$ErrorContacts
    )

    if (-not $enableEmailReport) {
        Write-Log "Email reporting is disabled. Skipping report." -Tag "Info"
        return
    }

    if ([string]::IsNullOrWhiteSpace($reportSmtpServer) -or
        [string]::IsNullOrWhiteSpace($reportFrom)       -or
        [string]::IsNullOrWhiteSpace($reportTo)) {

        Write-Log "Email report configuration incomplete (SMTP/From/To). Skipping report." -Tag "Error"
        return
    }

    $bodyLines = @()
    $bodyLines += "Contact sync report $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $bodyLines += ""
    $bodyLines += "Summary:"
    $bodyLines += "  Total:   $TotalContacts"
    $bodyLines += "  Created: $CreatedCount"
    $bodyLines += "  Updated: $UpdatedCount"
    $bodyLines += "  Skipped: $SkippedCount"
    $bodyLines += "  Deleted: $DeletedCount"
    $bodyLines += "  Errors:  $ErrorCount"
    $bodyLines += ""

    if ($CreatedContacts.Count -gt 0) {
        $bodyLines += "Created contacts:"
        foreach ($entry in $CreatedContacts) {
            $bodyLines += "  - $entry"
        }
        $bodyLines += ""
    }

    if ($UpdatedContacts.Count -gt 0) {
        $bodyLines += "Updated contacts:"
        foreach ($entry in $UpdatedContacts) {
            $bodyLines += "  - $entry"
        }
        $bodyLines += ""
    }

    if ($DeletedContacts.Count -gt 0) {
        $bodyLines += "Deleted contacts:"
        foreach ($entry in $DeletedContacts) {
            $bodyLines += "  - $entry"
        }
        $bodyLines += ""
    }

    if ($SkippedContacts.Count -gt 0) {
        $bodyLines += "Skipped contacts:"
        foreach ($entry in $SkippedContacts) {
            $bodyLines += "  - $entry"
        }
        $bodyLines += ""
    }

    if ($ErrorContacts.Count -gt 0) {
        $bodyLines += "Contacts with errors:"
        foreach ($entry in $ErrorContacts) {
            $bodyLines += "  - $entry"
        }
        $bodyLines += ""
    }

    $body = [string]::Join("`r`n", $bodyLines)

    try {
        Write-Log "Sending contact sync report to '$reportTo' via '$($reportSmtpServer):$($reportSmtpPort)'..." -Tag "Info"

        $sendMailParams = @{
            SmtpServer = $reportSmtpServer
            Port       = $reportSmtpPort
            From       = $reportFrom
            To         = $reportTo
            Subject    = $reportSubject
            Body       = $body
        }

        if ($reportUseSsl) {
            $sendMailParams["UseSsl"] = $true
        }

        if ($null -ne $reportSmtpCredential) {
            $sendMailParams["Credential"] = $reportSmtpCredential
        }

        Send-MailMessage @sendMailParams

        Write-Log "Contact sync report sent successfully." -Tag "Success"
    } catch {
        Write-Log "Failed to send contact sync report: $_" -Tag "Error"
    }
}

# ---------------------------[ Helper: Data Functions ]---------------------------

function Get-DirectoryOrgContacts {
    <#
        .SYNOPSIS
        Retrieves all organizational contacts (orgContact) from Microsoft Graph.
    #>
    try {
        Write-Log "Retrieving organizational contacts (orgContact) from Microsoft Graph..." -Tag "Debug"

        $orgContacts = Get-MgContact -All -Property `
            "id,displayName,givenName,surname,companyName,department,jobTitle,mail,mailNickname,addresses,phones,imAddresses"

        Write-Log "Retrieved $($orgContacts.Count) contacts from directory." -Tag "Success"
        return $orgContacts
    } catch {
        Write-Log "Failed to retrieve contacts: $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

function Get-MailboxContactIndex {
    <#
        .SYNOPSIS
        Retrieves all personal contacts from the target mailbox via Graph
        and returns a PSCustomObject with:
          - Index: hashtable indexed by primary email (lowercase)
          - Contacts: full contact list
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

        return [PSCustomObject]@{
            Index    = $contactIndex
            Contacts = $mailboxContacts
        }
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

                Write-Log "Compare businessPhones: Desired='$desiredString' Existing='$existingString'" -Tag "Debug"

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

                Write-Log "Compare imAddresses: Desired='$desiredString' Existing='$existingString'" -Tag "Debug"

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

                Write-Log "Compare emailAddresses: Desired='$desiredEmail' Existing='$existingEmail'" -Tag "Debug"

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

                Write-Log "Compare businessAddress: Desired='$desiredString' Existing='$existingString'" -Tag "Debug"

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

                Write-Log "Compare $property ($graphPropertyName): Desired='$desiredValue' Existing='$existingValue'" -Tag "Debug"

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

    Write-Log "Starting sync of Contacts into mailbox '$TargetMailboxUpn'..." -Tag "Info"

    $orgContacts           = Get-DirectoryOrgContacts
    $mailboxContactsResult = Get-MailboxContactIndex -UserId $TargetMailboxUpn
    $mailboxContactIndex   = $mailboxContactsResult.Index
    $mailboxContacts       = $mailboxContactsResult.Contacts

    $totalContacts = $orgContacts.Count
    $createdCount  = 0
    $updatedCount  = 0
    $skippedCount  = 0
    $deletedCount  = 0
    $errorCount    = 0

    $createdContactsList = @()
    $updatedContactsList = @()
    $skippedContactsList = @()
    $deletedContactsList = @()
    $errorContactsList   = @()

    # Build set of source emails from orgContacts
    $sourceEmailSet = @{}

    foreach ($orgContact in $orgContacts) {

        $primaryEmail = $orgContact.Mail
        if ([string]::IsNullOrWhiteSpace($primaryEmail)) {
            Write-Log "orgContact '$($orgContact.Id)' has no Mail property. Skipping." -Tag "Debug"
            $skippedCount++
            $skippedContactsList += "NO-EMAIL (Id: $($orgContact.Id))"
            continue
        }

        $emailKey = $primaryEmail.ToLowerInvariant()
        if (-not $sourceEmailSet.ContainsKey($emailKey)) {
            $sourceEmailSet[$emailKey] = $true
        }

        $desiredBody = Get-DesiredContactBody -OrgContact $orgContact

        if (-not $mailboxContactIndex.ContainsKey($emailKey)) {

            try {
                Write-Log "Creating new contact for '$primaryEmail' in mailbox '$TargetMailboxUpn'." -Tag "Info"
                New-MgUserContact -UserId $TargetMailboxUpn -BodyParameter $desiredBody | Out-Null
                $createdCount++
                $createdContactsList += $primaryEmail
            } catch {
                Write-Log "Failed to create contact for '$primaryEmail': $_" -Tag "Error"
                $errorCount++
                $errorContactsList += $primaryEmail
            }

        } else {

            $existingContact = $mailboxContactIndex[$emailKey]
            $updateBody      = Get-ContactUpdateBody -DesiredBody $desiredBody -ExistingContact $existingContact

            if ($updateBody.Keys.Count -eq 0) {
                Write-Log "Contact '$primaryEmail' already up to date. Skipping." -Tag "Debug"
                $skippedCount++
                $skippedContactsList += $primaryEmail
                continue
            }

            try {
                Write-Log "Updating contact '$primaryEmail' in mailbox '$TargetMailboxUpn' (fields: $($updateBody.Keys -join ', '))." -Tag "Info"
                Update-MgUserContact -UserId $TargetMailboxUpn -ContactId $existingContact.Id -BodyParameter $updateBody | Out-Null
                $updatedCount++
                $updatedContactsList += $primaryEmail
            } catch {
                Write-Log "Failed to update contact '$primaryEmail': $_" -Tag "Error"
                $errorCount++
                $errorContactsList += $primaryEmail
            }
        }
    }

    # Handle deletions: mailbox contacts that are no longer present as orgContacts
    if ($enableDeletion) {
        Write-Log "Checking for mailbox contacts that no longer exist as Contacts (deletions)..." -Tag "Info"

        foreach ($mailboxContact in $mailboxContacts) {
            if ($null -eq $mailboxContact.EmailAddresses -or $mailboxContact.EmailAddresses.Count -eq 0) {
                continue
            }

            $mbEmail = $mailboxContact.EmailAddresses[0].Address
            if ([string]::IsNullOrWhiteSpace($mbEmail)) {
                continue
            }

            $mbKey = $mbEmail.ToLowerInvariant()

            if (-not $sourceEmailSet.ContainsKey($mbKey)) {
                try {
                    Write-Log "Deleting mailbox contact '$mbEmail' because no matching orgContact exists." -Tag "Info"
                    Remove-MgUserContact -UserId $TargetMailboxUpn -ContactId $mailboxContact.Id -Confirm:$false
                    $deletedCount++
                    $deletedContactsList += $mbEmail
                } catch {
                    Write-Log "Failed to delete mailbox contact '$mbEmail': $_" -Tag "Error"
                    $errorCount++
                    $errorContactsList += $mbEmail
                }
            }
        }
    } else {
        Write-Log "Deletion of mailbox contacts is disabled (`$enableDeletion = `$false)." -Tag "Info"
    }

    Write-Log "Sync summary: Total=$totalContacts | Created=$createdCount | Updated=$updatedCount | Skipped=$skippedCount | Deleted=$deletedCount | Errors=$errorCount" -Tag "Success"

    # Send summary report via email
    Send-ContactSyncReport -TotalContacts $totalContacts `
                           -CreatedCount $createdCount `
                           -UpdatedCount $updatedCount `
                           -SkippedCount $skippedCount `
                           -DeletedCount $deletedCount `
                           -ErrorCount $errorCount `
                           -CreatedContacts $createdContactsList `
                           -UpdatedContacts $updatedContactsList `
                           -SkippedContacts $skippedContactsList `
                           -DeletedContacts $deletedContactsList `
                           -ErrorContacts $errorContactsList
}

# ---------------------------[ Main Execution ]---------------------------

Test-MicrosoftGraphConnection

Invoke-OrgContactSync -TargetMailboxUpn $targetMailboxUpn

Complete-Script -ExitCode 0
