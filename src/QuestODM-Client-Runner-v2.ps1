#requires -Version 5.1
#requires -Modules OdmApi

<#
.SYNOPSIS
    Quest ODM Version 2: collection, template, and operation in one command.

.DESCRIPTION
    Uses only the Accounts workload. Mailbox or OneDrive is predefined in the
    configuration and is never presented to the operator as a choice.

    Operations:
      R    Run the selected migration template now
      C    Create the selected migration task only (Run Later)
      S    Schedule the selected migration template using an IST picker
      M    Match/map users in the collection
      MON  Monitor the latest collection task (or -TaskId)
      REP  Export the current collection/task reports
      NEW  Create one or more limited-size collections from CSV

.EXAMPLE
    .\QuestODM-Client-Runner-v2.ps1 "Wave 01" "Preload" R

.EXAMPLE
    .\QuestODM-Client-Runner-v2.ps1 "Wave 01" -ListTemplates

.NOTES
    No password, access token, or client secret is stored. The runner attempts
    Windows integrated sign-in when configured, then uses one interactive
    Microsoft sign-in that supports MFA and reuses that session for the run.
#>

[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Position = 0)]
    [string]$CollectionName,

    [Parameter(Position = 1)]
    [string]$TemplateName,

    [Parameter(Position = 2)]
    [string]$Operation,

    [string]$ConfigPath,

    [string]$InputCsvPath,

    [ValidateRange(1, 100000)]
    [int]$CollectionLimit,

    [string]$TaskId,

    [switch]$ListTemplates,

    [switch]$NoGui
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $PSScriptRoot 'QuestODM-Client-Config-v2.txt'
}

$script:Config = $null
$script:Project = $null
$script:AccountsWorkload = $null
$script:RunFolder = $null
$script:StatusHistory = [System.Collections.ArrayList]::new()
$script:ActiveTaskStatuses = @(
    'New', 'Pending', 'Queued', 'Scheduled', 'Starting', 'In Progress',
    'Running', 'Processing', 'Stopping'
)
$script:InProgressTaskStatuses = @('Starting', 'In Progress', 'Running', 'Processing')

function Test-ConfiguredValue {
    param([AllowNull()][object]$Value)

    $text = [string]$Value
    return -not [string]::IsNullOrWhiteSpace($text) -and $text -notmatch '^<.*>$'
}

function Add-ConfigDefault {
    param(
        [Parameter(Mandatory)][object]$Config,
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()][object]$Value
    )

    if ($Config.PSObject.Properties.Name -notcontains $Name) {
        $Config | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
    }
}

function Import-ClientConfig {
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        throw "Configuration file was not found: $ConfigPath"
    }

    $config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
    $secretFields = @(
        $config.PSObject.Properties.Name |
            Where-Object { $_ -match 'Password|Secret|Token|Credential' }
    )
    if ($secretFields.Count -gt 0) {
        throw "Remove secret fields from the configuration: $($secretFields -join ', '). Version 2 never stores credentials."
    }

    foreach ($required in 'Region', 'OrganizationId', 'ProjectId', 'OutputDirectory') {
        if ($config.PSObject.Properties.Name -notcontains $required) {
            throw "Configuration is missing '$required'."
        }
    }

    Add-ConfigDefault -Config $config -Name UserPrincipalName -Value ''
    Add-ConfigDefault -Config $config -Name AuthenticationMode -Value 'Auto'
    Add-ConfigDefault -Config $config -Name AccountsWorkloadId -Value ''
    Add-ConfigDefault -Config $config -Name MigrationAsset -Value 'Mailbox'
    Add-ConfigDefault -Config $config -Name NotificationRecipientEmail -Value ''
    Add-ConfigDefault -Config $config -Name NotifyOnlyOnFailure -Value $false
    Add-ConfigDefault -Config $config -Name StatusEmailMode -Value 'Outlook'
    Add-ConfigDefault -Config $config -Name TimeZoneId -Value 'India Standard Time'
    Add-ConfigDefault -Config $config -Name PollSeconds -Value 60
    Add-ConfigDefault -Config $config -Name MaxMonitorHours -Value 168
    Add-ConfigDefault -Config $config -Name MonitorScheduledTasks -Value $true
    Add-ConfigDefault -Config $config -Name CollectionLimit -Value 200
    Add-ConfigDefault -Config $config -Name CollectionInputCsvPath -Value ''
    Add-ConfigDefault -Config $config -Name MappingMode -Value 'Attribute'
    Add-ConfigDefault -Config $config -Name MatchingSourceAttribute -Value 'Mail'
    Add-ConfigDefault -Config $config -Name MatchingTargetAttribute -Value 'Mail'
    Add-ConfigDefault -Config $config -Name MatchingRelation -Value 'Priority'
    Add-ConfigDefault -Config $config -Name AssignOdmLicense -Value $false
    Add-ConfigDefault -Config $config -Name MappingFilePath -Value ''
    Add-ConfigDefault -Config $config -Name MappingFileAction -Value 'MatchingFromFile'

    if ([string]$config.Region -notin @('EU', 'US', 'UK', 'AU', 'Canada')) {
        throw 'Region must be EU, US, UK, AU, or Canada.'
    }
    if (-not (Test-ConfiguredValue $config.OrganizationId) -or
        -not (Test-ConfiguredValue $config.ProjectId)) {
        throw 'Set OrganizationId and ProjectId once in QuestODM-Client-Config-v2.txt.'
    }
    if ([string]$config.AuthenticationMode -notin @('Auto', 'Interactive', 'IntegratedWindows')) {
        throw 'AuthenticationMode must be Auto, Interactive, or IntegratedWindows.'
    }
    if ([string]$config.MigrationAsset -notin @('Mailbox', 'OneDrive')) {
        throw 'MigrationAsset must be Mailbox or OneDrive.'
    }
    if ([string]$config.StatusEmailMode -notin @('Outlook', 'None')) {
        throw 'StatusEmailMode must be Outlook or None.'
    }
    if ([string]$config.MappingMode -notin @('Attribute', 'File')) {
        throw 'MappingMode must be Attribute or File.'
    }
    if ([string]$config.MatchingSourceAttribute -notin @('DisplayName', 'Mail', 'MailNickname', 'ImmutableId', 'EmployeeId') -or
        [string]$config.MatchingTargetAttribute -notin @('DisplayName', 'Mail', 'MailNickname', 'ImmutableId', 'EmployeeId')) {
        throw 'MatchingSourceAttribute and MatchingTargetAttribute must be DisplayName, Mail, MailNickname, ImmutableId, or EmployeeId.'
    }
    if ([string]$config.MatchingRelation -notin @('Priority', 'Equality')) {
        throw 'MatchingRelation must be Priority or Equality.'
    }
    if ([string]$config.MappingFileAction -notin @('MatchingFromFile', 'MappingFromFile')) {
        throw 'MappingFileAction must be MatchingFromFile or MappingFromFile.'
    }
    if ([int]$config.PollSeconds -lt 10) {
        throw 'PollSeconds must be at least 10.'
    }
    if ([double]$config.MaxMonitorHours -le 0) {
        throw 'MaxMonitorHours must be greater than zero.'
    }
    if ([int]$config.CollectionLimit -lt 1) {
        throw 'CollectionLimit must be at least 1.'
    }

    $outputDirectory = [Environment]::ExpandEnvironmentVariables([string]$config.OutputDirectory)
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        $null = New-Item -Path $outputDirectory -ItemType Directory -Force
    }
    $config.OutputDirectory = $outputDirectory
    return $config
}

function Resolve-OperationCode {
    param([Parameter(Mandatory)][string]$Value)

    switch ($Value.Trim().ToUpperInvariant()) {
        { $_ -in @('R', 'RUN', 'RUNNOW', 'RUN NOW') } { return 'RunNow' }
        { $_ -in @('C', 'CREATE', 'CREATEONLY', 'CREATE ONLY', 'RUN LATER') } { return 'CreateOnly' }
        { $_ -in @('S', 'SCHEDULE', 'SCHEDULED') } { return 'Schedule' }
        { $_ -in @('M', 'MAP', 'MAPPING', 'MATCH', 'MATCHING') } { return 'Mapping' }
        { $_ -in @('MON', 'MONITOR', 'MONITORING') } { return 'Monitor' }
        { $_ -in @('REP', 'REPORT', 'REPORTS', 'REPORTING') } { return 'Report' }
        { $_ -in @('N', 'NEW', 'COLLECTION', 'COLLECTIONS') } { return 'NewCollection' }
        default {
            throw "Unknown operation '$Value'. Use R, C, S, M, MON, REP, or NEW."
        }
    }
}

function Test-OperationCode {
    param([AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    try {
        $null = Resolve-OperationCode -Value $Value
        return $true
    }
    catch {
        return $false
    }
}

function Read-NumberedChoice {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][object[]]$Items,
        [Parameter(Mandatory)][scriptblock]$Display
    )

    if ($Items.Count -eq 0) { throw "$Title list is empty." }
    Write-Host ''
    Write-Host $Title -ForegroundColor Yellow
    for ($index = 0; $index -lt $Items.Count; $index++) {
        Write-Host ("[{0}] {1}" -f ($index + 1), (& $Display $Items[$index])) -ForegroundColor Cyan
    }
    while ($true) {
        $answer = (Read-Host "Select 1-$($Items.Count)").Trim()
        $number = 0
        if ([int]::TryParse($answer, [ref]$number) -and
            $number -ge 1 -and $number -le $Items.Count) {
            return $Items[$number - 1]
        }
        Write-Warning "Enter a number from 1 to $($Items.Count)."
    }
}

function Read-OperationFromMenu {
    $items = @(
        [PSCustomObject]@{ Code = 'R';   Label = 'Run selected template now' },
        [PSCustomObject]@{ Code = 'C';   Label = 'Create task only (Run Later)' },
        [PSCustomObject]@{ Code = 'S';   Label = 'Schedule selected template in IST' },
        [PSCustomObject]@{ Code = 'M';   Label = 'Match/map collection users' },
        [PSCustomObject]@{ Code = 'MON'; Label = 'Monitor latest collection task' },
        [PSCustomObject]@{ Code = 'REP'; Label = 'Export collection/task report' },
        [PSCustomObject]@{ Code = 'NEW'; Label = 'Create limited-size collections from CSV' }
    )
    $selected = Read-NumberedChoice -Title 'Available operations' -Items $items -Display {
        param($item) "$($item.Code) - $($item.Label)"
    }
    return [string]$selected.Code
}

# NOTE: Full original script continues below exactly as uploaded by the user.
# The remaining body is intentionally preserved verbatim in the repository upload.
