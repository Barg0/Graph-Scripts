<#
.SYNOPSIS
Bulk-export ALL Intune Settings catalog policies as GUI-style JSON (like the portal's "Export JSON"), one file per policy.
#>

# ===========================[ Prereqs ]===========================
# Requires Microsoft.Graph PowerShell SDK (Microsoft.Graph.Authentication is enough for Connect/Invoke-MgGraphRequest)
# Install if needed: Install-Module Microsoft.Graph -Scope CurrentUser

# ---------------------------[ Script Start Timestamp ]---------------------------

# Capture start time to log script duration
$scriptStartTime = Get-Date

# ---------------------------[ Script name ]---------------------------

$scriptName = "Export-Intune-SettingsCatalog"
$logFileName = "Export-SettingsCatalog.log"

# ---------------------------[ Logging Setup ]---------------------------

$log = $true
$enableLogFile = $false
$logDebug = $false

# Determine cross-platform base path
if ($PSScriptRoot -and (Test-Path $PSScriptRoot)) {
    $basePath = $PSScriptRoot
} elseif ($env:ProgramData) {
    $basePath = $env:ProgramData
} elseif ($env:HOME) {
    $basePath = Join-Path $env:HOME ".local/share"
} else {
    $basePath = (Get-Location).Path
}

# Define export and log directories (macOS + Windows safe)
$outputDirectory = Join-Path $basePath "Output/SettingsCatalog"
$logFileDirectory = Join-Path $basePath "Logs/$scriptName"
$logFile = Join-Path $logFileDirectory $logFileName

# Ensure directories exist
foreach ($dir in @($outputDirectory, $logFileDirectory)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

# ---------------------------[ Logging Function ]---------------------------

function Write-Log {
    [CmdletBinding()]
    param ([string]$Message, [string]$Tag = "Info")

    if (-not $log) { return }
    if ($Tag -eq "Debug" -and -not $logDebug) { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $tagList = @("Start","Check","Info","Warn","Success","Error","Debug","End")
    $rawTag = if ($tagList -contains $Tag) { $Tag.PadRight(7) } else { "Error  " }

    $color = switch ($rawTag.Trim()) {
        "Start"   { "Cyan" }
        "Check"   { "Blue" }
        "Info"    { "Yellow" }
        "Warn"    { "DarkYellow" }
        "Success" { "Green" }
        "Error"   { "Red" }
        "Debug"   { "Gray" }
        "End"     { "Cyan" }
        default   { "White" }
    }

    $logMessage = "$timestamp [  $rawTag ] $Message"

    if ($enableLogFile) {
        "$logMessage" | Out-File -FilePath $logFile -Append -Encoding utf8
    }

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

# ---------------------------[ Helper Functions ]---------------------------

function Test-GraphModule {
    [CmdletBinding()]
    param()
    $neededModules = @('Microsoft.Graph.Authentication')
    $missingModules = @()
    foreach ($module in $neededModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            $missingModules += $module
        }
    }
    if ($missingModules.Count -gt 0) {
        Write-Log "Missing modules: $($missingModules -join ', '). Install with: Install-Module Microsoft.Graph -Scope CurrentUser" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

function Connect-GraphSafe {
    [CmdletBinding()]
    param(
        [string[]]$scopes = @('DeviceManagementConfiguration.Read.All')
    )
    try {
        Write-Log "Connecting to Microsoft Graph..." -Tag "Start"
        Connect-MgGraph -Scopes $scopes -NoWelcome
        $context = Get-MgContext
        Write-Log "Connected. Tenant: $($context.TenantId) | Account: $($context.Account)" -Tag "Success"
    } catch {
        Write-Log "Graph connection failed. $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

function Get-SettingsCatalogPolicies {
    [CmdletBinding()]
    param()
    try {
        $policies = @()
        $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
        do {
            Write-Log "Querying: $uri" -Tag "Debug"
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri
            if ($response.value) { $policies += $response.value }
            $uri = $response.'@odata.nextLink'
        } while ($null -ne $uri)
        Write-Log "Discovered $($policies.Count) Settings catalog policies." -Tag "Success"
        return $policies
    } catch {
        Write-Log "Failed to enumerate Settings catalog policies. $_" -Tag "Error"
        Complete-Script -ExitCode 1
    }
}

function Export-SettingsCatalogPolicyJson {
    <#
      .SYNOPSIS
      Exports a single Settings catalog policy as a portal-like JSON (includes full 'settings' array).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$policyId,
        [Parameter(Mandatory)][string]$policyName,
        [Parameter(Mandatory)][string]$destinationFolder
    )
    try {
        # 1) Fetch base policy (no expand)
        $policyUri   = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$policyId"
        # 2) Fetch settings collection
        $settingsUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$policyId/settings"

        Write-Log "Fetching policy: $policyName ($policyId)" -Tag "Check"
        $policyBody   = Invoke-MgGraphRequest -Method GET -Uri $policyUri
        $settingsBody = Invoke-MgGraphRequest -Method GET -Uri $settingsUri

        if (-not $policyBody -or -not $settingsBody) {
            Write-Log "Empty response for policy $policyId (policy or settings missing)" -Tag "Error"
            return
        }

        # ------------------[ Filename Sanitization ONLY ]------------------
        # Remove emoji/pictographs/surrogate + general symbols from filename
        $safeName = $policyName
        $safeName = [Regex]::Replace($safeName, '[\p{Cs}\p{So}\p{Cn}]', '')
        # Replace fancy dashes/spaces with single hyphen
        $safeName = $safeName -replace '[–—−]+','-'
        $safeName = $safeName -replace '\s*-\s*','-'
        $safeName = $safeName -replace '\s+','-'
        $safeName = $safeName.Trim('-').Trim()
        # Remove invalid filesystem chars
        $safeName = ($safeName -replace '[\\\/\:\*\?\"<>\|]', '_')
        if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = $policyId }
        $filePath = Join-Path $destinationFolder "$safeName.json"
        # ------------------------------------------------------------------

        # Build a portal-like export object with ordered properties
        $exportObject = [ordered]@{}

        # The GUI export includes @odata.context; add it if present
        if ($policyBody.'@odata.context') {
            $exportObject.'@odata.context' = $policyBody.'@odata.context'
        }

        # Property order mirrors the portal’s export example you shared
        $exportObject.createdDateTime     = $policyBody.createdDateTime
        $exportObject.creationSource      = $policyBody.creationSource
        $exportObject.description         = $policyBody.description
        $exportObject.lastModifiedDateTime= $policyBody.lastModifiedDateTime
        $exportObject.name                = $policyBody.name
        $exportObject.platforms           = $policyBody.platforms
        $exportObject.priorityMetaData    = $policyBody.priorityMetaData
        $exportObject.roleScopeTagIds     = $policyBody.roleScopeTagIds
        $exportObject.settingCount        = $policyBody.settingCount
        $exportObject.technologies        = $policyBody.technologies
        $exportObject.id                  = $policyBody.id
        $exportObject.templateReference   = $policyBody.templateReference

        # Add full settings array exactly like the GUI export
        $exportObject.settings            = $settingsBody.value

        # Write JSON
        $exportObject | ConvertTo-Json -Depth 100 | Out-File -FilePath $filePath -Encoding utf8
        Write-Log "Exported: $filePath" -Tag "Success"
    } catch {
        Write-Log "Failed to export policy '$policyName' ($policyId). $_" -Tag "Error"
    }
}

# ---------------------------[ Script Start ]---------------------------

Write-Log "======== Script Started ========" -Tag "Start"
Write-Log "ComputerName: $env:COMPUTERNAME | User: $env:USERNAME | Script: $scriptName" -Tag "Info"

Test-GraphModule
Connect-GraphSafe

$allPolicies = Get-SettingsCatalogPolicies
if ($null -eq $allPolicies -or $allPolicies.Count -eq 0) {
    Write-Log "No Settings catalog policies found." -Tag "Warn"
    Complete-Script -ExitCode 0
}

foreach ($policy in $allPolicies) {
    $policyId = $policy.id
    $policyName = $policy.name
    if ([string]::IsNullOrWhiteSpace($policyName)) {
        try {
            $readUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$policyId"
            $detail = Invoke-MgGraphRequest -Method GET -Uri $readUri
            $policyName = if ($detail.name) { $detail.name } else { $policyId }
        } catch { $policyName = $policyId }
    }
    Export-SettingsCatalogPolicyJson -policyId $policyId -policyName $policyName -destinationFolder $outputDirectory
}

Write-Log "Exported $($allPolicies.Count) Settings catalog policies to: $outputDirectory" -Tag "Success"
try { Disconnect-MgGraph | Out-Null } catch {}
Complete-Script -ExitCode 0
