# Script version:   2025-06-07 17:15
# Script author:    Barg0

$scriptStartTime = Get-Date

# ---------------------------[ Parameters ]---------------------------

$graphScopes = @(
    "DeviceManagementConfiguration.Read.All",
    "DeviceManagementConfiguration.ReadWrite.All",
    "Group.Read.All"
)
$scriptName = "Assign-IntunePolicies"
$logFileName = "$($scriptName).log"

$csvPath = Join-Path $PSScriptRoot -ChildPath "PolicyAssignments.csv"

# ---------------------------[ Logging Control ]---------------------------

$log = $true
$enableLogFile = $false
$logFileDirectory = "$PSScriptRoot"
$logFile = Join-Path -Path $logFileDirectory -ChildPath $logFileName

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
    if ($enableLogFile) { "$logMessage" | Out-File -FilePath $logFile -Append }

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
    if ($null -ne $context.Account -and $null -ne $context.Scopes -and ($graphScopes | ForEach-Object { $context.Scopes -contains $_ }) -contains $true) {
        Write-Log "Microsoft Graph already connected as $($context.Account)" -Tag "Success"
        $connected = $true
    } else {
        Write-Log "Microsoft Graph context incomplete or lacks required scopes. Reconnecting..." -Tag "Info"
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

# ---------------------------[ Helper Functions ]---------------------------

function Find-Policy {
    param([string]$PolicyName)

    $configPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -ErrorAction Stop
    $match = $configPolicies.value | Where-Object { $_.name -eq $PolicyName }
    if ($match) {
        return @{
            Type = "ConfigurationPolicy"
            Id = $match.id
            Uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($match.id)/assign"
            AssignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($match.id)/assignments"
        }
    }

    $deviceConfigs = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations" -ErrorAction Stop
    $match = $deviceConfigs.value | Where-Object { $_.displayName -eq $PolicyName }
    if ($match) {
        return @{
            Type = "CSP"
            Id = $match.id
            Uri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($match.id)/assign"
            AssignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($match.id)/groupAssignments"
        }
    }

    return $null
}

function Resolve-GroupId {
    param (
        [string]$GroupName
    )

    if ($GroupName -eq "All Users" -or $GroupName -eq "All Devices") {
        return $null
    }

    try {
        $groupQuery = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$GroupName'" -ErrorAction Stop
        $group = $groupQuery.value | Where-Object { $_.securityEnabled -eq $true }
        if ($group) {
            return $group.id
        } else {
            Write-Log "Group '$GroupName' not found or not security-enabled." -Tag "Error"
            return $null
        }
    } catch {
        Write-Log "Error resolving group '$GroupName': $_" -Tag "Error"
        return $null
    }
}

function Set-IntunePolicyAssignment {
    param(
        [string]$PolicyName,
        [string[]]$IncludeGroups,
        [string[]]$ExcludeGroups
    )

    $policy = Find-Policy -PolicyName $PolicyName
    if ($null -eq $policy) {
        Write-Log "Policy '$PolicyName' not found in Intune" -Tag "Error"
        return
    }

    $assignments = @()

    foreach ($g in $IncludeGroups) {
        $g = $g.Trim()
        if ($g) {
            if ($g -eq "All Users") {
                $assignments += @{
                    target = @{
                        "@odata.type" = "#microsoft.graph.allLicensedUsersAssignmentTarget"
                    }
                }
            } elseif ($g -eq "All Devices") {
                $assignments += @{
                    target = @{
                        "@odata.type" = "#microsoft.graph.allDevicesAssignmentTarget"
                    }
                }
            } else {
                $groupId = Resolve-GroupId -GroupName $g
                if ($groupId) {
                    $assignments += @{
                        target = @{
                            "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                            groupId = $groupId
                        }
                    }
                }
            }
        }
    }

    foreach ($g in $ExcludeGroups) {
        $g = $g.Trim()
        if ($g) {
            $groupId = Resolve-GroupId -GroupName $g
            if ($groupId) {
                $assignments += @{
                    target = @{
                        "@odata.type" = "#microsoft.graph.exclusionGroupAssignmentTarget"
                        groupId = $groupId
                    }
                }
            }
        }
    }

    if ($assignments.Count -eq 0) {
        Write-Log "No valid targets found for policy '$PolicyName'. Skipping." -Tag "Error"
        return
    }

    $body = @{ assignments = $assignments } | ConvertTo-Json -Depth 10

    try {
        Invoke-MgGraphRequest -Method POST -Uri $policy.Uri -Body $body -ErrorAction Stop
        Write-Log "Policy '$PolicyName' assigned successfully with include/exclude targets." -Tag "Success"
    } catch {
        Write-Log "Failed to assign policy '$PolicyName': $_" -Tag "Error"
    }
}

# ---------------------------[ CSV Processing ]---------------------------

if (-not (Test-Path $csvPath)) {
    Write-Log "CSV file not found: $csvPath" -Tag "Error"
    Complete-Script -ExitCode 1
}

try {
    $rows = Import-Csv -Path $csvPath
    foreach ($entry in $rows) {
        $policyName = $entry.Policy.Trim()
        $include = if ($entry.Include) { $entry.Include -split ';' } else { @() }
        $exclude = if ($entry.Exclude) { $entry.Exclude -split ';' } else { @() }

        Set-IntunePolicyAssignment -PolicyName $policyName -IncludeGroups $include -ExcludeGroups $exclude
    }
} catch {
    Write-Log "Failed to process CSV: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

Complete-Script -ExitCode 0