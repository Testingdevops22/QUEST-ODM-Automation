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

function Get-WorkloadType {
    param([Parameter(Mandatory)][object]$Workload)

    foreach ($name in 'WorkloadType', 'Type', 'Name') {
        if ($Workload.PSObject.Properties.Name -contains $name -and
            -not [string]::IsNullOrWhiteSpace([string]$Workload.$name)) {
            return [string]$Workload.$name
        }
    }
    return ''
}

function Assert-QuestCommands {
    $required = @(
        'Connect-OdmService', 'Select-OdmOrganization', 'Get-OdmProject',
        'Get-OdmProjectWorkload', 'Select-OdmProjectWorkload', 'Get-OdmCollection',
        'New-OdmCollection', 'Get-OdmObject', 'Get-OdmTask', 'Get-OdmEvent',
        'Get-OdmTaskTemplates', 'New-OdmMatchingTask', 'New-OdmMappingFileTask',
        'New-OdmMailMigrationTask', 'New-OdmOneDriveMigrationTask',
        'Add-OdmObject', 'Add-OdmTask', 'Start-OdmTask'
    )
    $missing = @($required | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($missing.Count -gt 0) {
        throw "The installed OdmApi module is missing: $($missing -join ', '). Update the module and retest Version 2."
    }
}

function Connect-Quest {
    if (-not (Get-Module OdmApi)) {
        Import-Module OdmApi -ErrorAction Stop
    }
    Assert-QuestCommands

    $moduleVersion = (Get-Module OdmApi | Sort-Object Version -Descending | Select-Object -First 1).Version
    Write-Host "Quest OdmApi $moduleVersion" -ForegroundColor DarkGray

    $mode = [string]$script:Config.AuthenticationMode
    $runningOnWindows = $env:OS -eq 'Windows_NT'
    $connected = $false

    if ($mode -in @('Auto', 'IntegratedWindows') -and $runningOnWindows) {
        try {
            Write-Host 'Trying Windows single sign-on...' -ForegroundColor Cyan
            Connect-OdmService -Region ([string]$script:Config.Region) -UseIntegratedWindowsAuth
            $connected = $true
        }
        catch {
            if ($mode -eq 'IntegratedWindows') { throw }
            Write-Warning "Windows single sign-on was unavailable: $($_.Exception.Message)"
        }
    }

    if (-not $connected) {
        if (Test-ConfiguredValue $script:Config.UserPrincipalName) {
            Write-Host "Complete the Microsoft MFA sign-in for $($script:Config.UserPrincipalName)." -ForegroundColor Cyan
        }
        else {
            Write-Host 'Complete the Microsoft MFA sign-in.' -ForegroundColor Cyan
        }
        Connect-OdmService -Region ([string]$script:Config.Region)
    }

    Select-OdmOrganization -OrganizationId ([string]$script:Config.OrganizationId)
    $projects = @(Get-OdmProject -ProjectId ([string]$script:Config.ProjectId))
    if ($projects.Count -ne 1) {
        throw "Expected one Quest project '$($script:Config.ProjectId)' but found $($projects.Count)."
    }
    $script:Project = $projects[0]

    $workloads = @(Get-OdmProjectWorkload -Project $script:Project -All)
    $configuredWorkloadId = [string]$script:Config.AccountsWorkloadId
    if (Test-ConfiguredValue $configuredWorkloadId) {
        $accounts = @($workloads | Where-Object { [string]$_.Id -eq $configuredWorkloadId })
    }
    else {
        $accounts = @($workloads | Where-Object { (Get-WorkloadType $_) -eq 'Accounts' })
    }
    if ($accounts.Count -ne 1) {
        throw "Expected one Accounts workload but found $($accounts.Count). Set AccountsWorkloadId if the project has more than one."
    }

    $script:AccountsWorkload = $accounts[0]
    Select-OdmProjectWorkload -ProjectWorkloadId ([string]$script:AccountsWorkload.Id)
    Write-Host "Connected once. Project: $($script:Project.Name)" -ForegroundColor Green
}

function Get-SafeFileName {
    param([Parameter(Mandatory)][string]$Value)

    $safe = $Value
    foreach ($character in [IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$character, '_')
    }
    if ($safe.Length -gt 80) { $safe = $safe.Substring(0, 80) }
    return $safe
}

function New-RunFolder {
    param(
        [Parameter(Mandatory)][string]$Collection,
        [Parameter(Mandatory)][string]$Action
    )

    $folderName = '{0}-{1}-{2}' -f (
        Get-Date -Format 'yyyyMMdd-HHmmss'
    ), (Get-SafeFileName $Collection), (Get-SafeFileName $Action)
    $script:RunFolder = Join-Path ([string]$script:Config.OutputDirectory) $folderName
    $null = New-Item -Path $script:RunFolder -ItemType Directory -Force
    return $script:RunFolder
}

function Find-AccountsCollection {
    param([Parameter(Mandatory)][string]$Name)

    $all = @(Get-OdmCollection -ResultSize 10000)
    $exact = @($all | Where-Object { [string]$_.Name -eq $Name })
    if ($exact.Count -eq 1) { return $exact[0] }
    if ($exact.Count -gt 1) { throw "More than one Accounts collection is named '$Name'." }

    $partial = @(
        $all | Where-Object {
            ([string]$_.Name).IndexOf($Name, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
        }
    )
    if ($partial.Count -eq 1) {
        Write-Host "Using Accounts collection '$($partial[0].Name)'." -ForegroundColor Green
        return $partial[0]
    }
    if ($partial.Count -gt 1) {
        throw "Collection text '$Name' matches $($partial.Count) Accounts collections. Enter the full collection name."
    }
    throw "Accounts collection '$Name' was not found."
}

function Select-AccountsCollectionFromMenu {
    $collections = @(Get-OdmCollection -ResultSize 10000 | Sort-Object Name)
    if ($collections.Count -eq 0) { throw 'No Accounts collections were found.' }
    return Read-NumberedChoice -Title 'Available Accounts collections' -Items $collections -Display {
        param($item) [string]$item.Name
    }
}

function Get-CollectionUsers {
    param([Parameter(Mandatory)][object]$Collection)

    $objects = @(
        Get-OdmObject -FilterObject $Collection -All -IncludeAllProperties -IncludeCollections
    )
    $users = @($objects | Where-Object { [string]$_.Type -eq 'User' })
    $ignored = $objects.Count - $users.Count
    if ($ignored -gt 0) {
        Write-Host "Ignored $ignored non-user object(s). Accounts users only." -ForegroundColor DarkGray
    }
    if ($users.Count -eq 0) {
        throw "Accounts collection '$($Collection.Name)' contains no user accounts."
    }
    return $users
}

function Get-PropertyValue {
    param(
        [Parameter(Mandatory)][object]$Object,
        [Parameter(Mandatory)][string]$Name
    )

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Get-CollectionNames {
    param([Parameter(Mandatory)][object]$Object)

    $collections = @(Get-PropertyValue -Object $Object -Name 'Collections')
    if ($collections.Count -eq 0) { return '' }
    return (@($collections | ForEach-Object {
        if ($_.PSObject.Properties.Name -contains 'Name') { [string]$_.Name } else { [string]$_ }
    }) -join '; ')
}

function ConvertTo-MailboxRows {
    param([Parameter(Mandatory)][object[]]$Users)

    return @($Users | ForEach-Object {
        [PSCustomObject][ordered]@{
            'Type'                       = Get-PropertyValue $_ 'RecipientTypeDetails'
            'Mailbox State'              = Get-PropertyValue $_ 'MailboxStatus'
            'Source Mailbox'             = Get-PropertyValue $_ 'SourceEmail'
            'Source UPN'                 = Get-PropertyValue $_ 'SourceUserPrincipalName'
            'Target Mailbox'             = Get-PropertyValue $_ 'TargetEmail'
            'Target UPN'                 = Get-PropertyValue $_ 'TargetUserPrincipalName'
            'ODM Licensed'               = Get-PropertyValue $_ 'OdmLicensed'
            'All items'                  = Get-PropertyValue $_ 'MailboxItemsCount'
            'Total Size (bytes)'         = Get-PropertyValue $_ 'MailboxSize'
            'Items to Migrate'           = Get-PropertyValue $_ 'MailMigrationEstimated'
            'Processed'                  = Get-PropertyValue $_ 'MailMigrationProcessed'
            'Errors'                     = Get-PropertyValue $_ 'MailMigrationErrors'
            'Archive Mailbox State'      = Get-PropertyValue $_ 'ArchiveMailboxStatus'
            'Archive Items to Migrate'   = Get-PropertyValue $_ 'ArchiveMailMigrationEstimated'
            'Archive Processed'          = Get-PropertyValue $_ 'ArchiveMailMigrationProcessed'
            'Collections'                = Get-CollectionNames $_
        }
    })
}

function ConvertTo-OneDriveRows {
    param([Parameter(Mandatory)][object[]]$Users)

    return @($Users | ForEach-Object {
        [PSCustomObject][ordered]@{
            'Name'                 = Get-PropertyValue $_ 'DisplayName'
            'Migration State'      = Get-PropertyValue $_ 'OneDriveMigrationState'
            'Collect Statistics'   = Get-PropertyValue $_ 'OneDriveAssessmentStatus'
            'Provisioning'         = Get-PropertyValue $_ 'OneDriveProvisioned'
            'Items Migrated (%)'   = Get-PropertyValue $_ 'OneDriveMigrationProgress'
            'Source Items'         = Get-PropertyValue $_ 'OneDriveSourceItemCount'
            'Source Size (bytes)'  = Get-PropertyValue $_ 'OneDriveSourceTotalSize'
            'Source Last Modified' = Get-PropertyValue $_ 'OneDriveSourceLastModified'
            'Target items'         = Get-PropertyValue $_ 'OneDriveTargetItemCount'
            'Last Successful Run'  = Get-PropertyValue $_ 'OneDriveLastSuccessfulRun'
            'Last Run Status'      = Get-PropertyValue $_ 'OneDriveStatus'
            'Source UPN'           = Get-PropertyValue $_ 'SourceUserPrincipalName'
            'Source Root Url'      = Get-PropertyValue $_ 'OneDriveSourceRootUrl'
            'Target UPN'           = Get-PropertyValue $_ 'TargetUserPrincipalName'
            'Collections'          = Get-CollectionNames $_
        }
    })
}

function Export-CollectionSnapshot {
    param(
        [Parameter(Mandatory)][object]$Collection,
        [Parameter(Mandatory)][ValidateSet('Pre', 'Post', 'Current', 'Post-Created')][string]$Phase
    )

    $users = Get-CollectionUsers -Collection $Collection
    $asset = [string]$script:Config.MigrationAsset
    $rows = if ($asset -eq 'Mailbox') {
        ConvertTo-MailboxRows -Users $users
    }
    else {
        ConvertTo-OneDriveRows -Users $users
    }
    $path = Join-Path $script:RunFolder "$Phase-$asset.csv"
    $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    Write-Host "$Phase report: $path" -ForegroundColor DarkGray
    return $path
}

function Export-TaskDetails {
    param([Parameter(Mandatory)][object]$Task)

    $fresh = @(Get-OdmTask -TaskId ([string]$Task.Id)) | Select-Object -First 1
    if ($null -eq $fresh) { $fresh = $Task }

    @([PSCustomObject][ordered]@{
        TaskId          = Get-PropertyValue $fresh 'Id'
        TaskName        = Get-PropertyValue $fresh 'Name'
        TaskType        = Get-PropertyValue $fresh 'Type'
        Status          = Get-PropertyValue $fresh 'Status'
        Progress        = Get-PropertyValue $fresh 'Progress'
        LastResult      = Get-PropertyValue $fresh 'LastResult'
        Created         = Get-PropertyValue $fresh 'Created'
        Modified        = Get-PropertyValue $fresh 'Modified'
        Collection      = $CollectionName
        ReportGenerated = Get-Date
    }) | Export-Csv -LiteralPath (Join-Path $script:RunFolder 'Task-Summary.csv') -NoTypeInformation -Encoding UTF8

    $events = @(Get-OdmEvent -FilterObject $fresh -All -ErrorAction SilentlyContinue)
    $eventRows = @($events | Select-Object Id, TaskId, TaskName, Timestamp, Severity, Category, Message, Details, ObjectId, ObjectName)
    if ($eventRows.Count -gt 0) {
        $eventRows | Export-Csv -LiteralPath (Join-Path $script:RunFolder 'Task-Events.csv') -NoTypeInformation -Encoding UTF8
    }
    else {
        '"Id","TaskId","TaskName","Timestamp","Severity","Category","Message","Details","ObjectId","ObjectName"' |
            Set-Content -LiteralPath (Join-Path $script:RunFolder 'Task-Events.csv') -Encoding UTF8
    }
}

function Export-StatusHistory {
    if ($script:StatusHistory.Count -gt 0) {
        $script:StatusHistory |
            Export-Csv -LiteralPath (Join-Path $script:RunFolder 'Status-History.csv') -NoTypeInformation -Encoding UTF8
    }
}

function Send-TaskStatus {
    param(
        [Parameter(Mandatory)][string]$Stage,
        [Parameter(Mandatory)][object]$Task,
        [Parameter(Mandatory)][object]$Collection,
        [string]$Detail
    )

    $entry = [PSCustomObject][ordered]@{
        Time       = Get-Date
        Stage      = $Stage
        TaskId     = Get-PropertyValue $Task 'Id'
        TaskName   = Get-PropertyValue $Task 'Name'
        TaskStatus = Get-PropertyValue $Task 'Status'
        Progress   = Get-PropertyValue $Task 'Progress'
        Collection = [string]$Collection.Name
        Detail     = $Detail
    }
    $null = $script:StatusHistory.Add($entry)
    Export-StatusHistory

    $recipient = [string]$script:Config.NotificationRecipientEmail
    if ([string]$script:Config.StatusEmailMode -eq 'None' -or
        [string]::IsNullOrWhiteSpace($recipient)) {
        return
    }

    if ($env:OS -ne 'Windows_NT') {
        Write-Warning "Status email '$Stage' was not sent: Outlook desktop automation requires Windows."
        return
    }

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.To = $recipient
        $mail.Subject = "Quest ODM: $Stage - $($Task.Name)"
        $mail.Body = @"
Quest ODM task update

Stage: $Stage
Task: $($Task.Name)
Task ID: $($Task.Id)
Collection: $($Collection.Name)
Status: $($Task.Status)
Progress: $($Task.Progress)
Details: $Detail
Report folder: $script:RunFolder

Sent by QuestODM Client Runner Version 2 using the signed-in Outlook profile.
"@
        $mail.Send()
        Write-Host "Email sent: $Stage" -ForegroundColor DarkGray
    }
    catch {
        Write-Warning "Status email '$Stage' could not be sent: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $mail) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($mail) }
        if ($null -ne $outlook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) }
    }
}

function Get-ConfiguredTimeZone {
    $ids = @([string]$script:Config.TimeZoneId, 'India Standard Time', 'Asia/Kolkata') | Select-Object -Unique
    foreach ($id in $ids) {
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        try { return [TimeZoneInfo]::FindSystemTimeZoneById($id) } catch { }
    }
    throw "IST time zone '$($script:Config.TimeZoneId)' is unavailable."
}

function Convert-IstToUtc {
    param([Parameter(Mandatory)][datetime]$IstTime)

    $zone = Get-ConfiguredTimeZone
    $unspecified = [datetime]::SpecifyKind($IstTime, [DateTimeKind]::Unspecified)
    return [TimeZoneInfo]::ConvertTimeToUtc($unspecified, $zone)
}

function Read-ScheduledTimeIst {
    $zone = Get-ConfiguredTimeZone
    $nowIst = [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $zone)
    $minimum = $nowIst.AddMinutes(10)
    $runningOnWindows = $env:OS -eq 'Windows_NT'

    if ($runningOnWindows -and -not $NoGui) {
        try {
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing

            $form = New-Object System.Windows.Forms.Form
            $form.Text = 'Schedule Quest ODM task (IST)'
            $form.Size = New-Object System.Drawing.Size(440, 175)
            $form.StartPosition = 'CenterScreen'
            $form.TopMost = $true

            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(20, 20)
            $label.Size = New-Object System.Drawing.Size(390, 25)
            $label.Text = 'Select the task start date and time in India Standard Time (IST):'
            $form.Controls.Add($label)

            $picker = New-Object System.Windows.Forms.DateTimePicker
            $picker.Location = New-Object System.Drawing.Point(20, 52)
            $picker.Size = New-Object System.Drawing.Size(385, 25)
            $picker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
            $picker.CustomFormat = 'yyyy-MM-dd HH:mm'
            $picker.MinDate = $minimum
            $picker.Value = $minimum.AddMinutes(5)
            $form.Controls.Add($picker)

            $ok = New-Object System.Windows.Forms.Button
            $ok.Location = New-Object System.Drawing.Point(245, 92)
            $ok.Size = New-Object System.Drawing.Size(75, 28)
            $ok.Text = 'Schedule'
            $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.AcceptButton = $ok
            $form.Controls.Add($ok)

            $cancel = New-Object System.Windows.Forms.Button
            $cancel.Location = New-Object System.Drawing.Point(330, 92)
            $cancel.Size = New-Object System.Drawing.Size(75, 28)
            $cancel.Text = 'Cancel'
            $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.CancelButton = $cancel
            $form.Controls.Add($cancel)

            if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
                throw 'Scheduling was cancelled. No task was created.'
            }
            return Convert-IstToUtc -IstTime $picker.Value
        }
        finally {
            if ($null -ne $form) { $form.Dispose() }
        }
    }

    while ($true) {
        $answer = (Read-Host 'Schedule time in IST (yyyy-MM-dd HH:mm)').Trim()
        $parsed = [datetime]::MinValue
        $valid = [datetime]::TryParseExact(
            $answer,
            'yyyy-MM-dd HH:mm',
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None,
            [ref]$parsed
        )
        if (-not $valid) {
            Write-Warning 'Use yyyy-MM-dd HH:mm, for example 2026-09-02 22:30.'
            continue
        }
        if ($parsed -lt $minimum) {
            Write-Warning "Choose a time at least 10 minutes ahead. Earliest: $($minimum.ToString('yyyy-MM-dd HH:mm')) IST"
            continue
        }
        return Convert-IstToUtc -IstTime $parsed
    }
}

function Get-AvailableTemplates {
    $templateType = if ([string]$script:Config.MigrationAsset -eq 'Mailbox') {
        'Mail Migration'
    }
    else {
        'OneDrive Migration'
    }

    return @(
        Get-OdmTaskTemplates -ProjectId ([string]$script:AccountsWorkload.Id) -All |
            Where-Object { [string]$_.TemplateType -eq $templateType } |
            Sort-Object Name
    )
}

function ConvertTo-CommandArgument {
    param([Parameter(Mandatory)][string]$Value)

    return '"{0}"' -f $Value.Replace('"', '`"')
}

function Show-AvailableTemplates {
    param([string]$CollectionForCommand)

    $templates = @(Get-AvailableTemplates)
    if ($templates.Count -eq 0) {
        throw "No saved $($script:Config.MigrationAsset) templates were found in the Accounts workload."
    }

    Write-Host ''
    Write-Host "Available $($script:Config.MigrationAsset) templates" -ForegroundColor Yellow
    for ($index = 0; $index -lt $templates.Count; $index++) {
        Write-Host ("[{0}] {1} | ID: {2}" -f ($index + 1), $templates[$index].Name, $templates[$index].Id) -ForegroundColor Cyan
    }

    $collectionText = if ([string]::IsNullOrWhiteSpace($CollectionForCommand)) {
        '"<collection name>"'
    }
    else {
        ConvertTo-CommandArgument -Value $CollectionForCommand
    }
    Write-Host ''
    Write-Host 'Ready command examples (change the final R to C or S when needed):' -ForegroundColor Green
    foreach ($template in $templates) {
        $templateText = ConvertTo-CommandArgument -Value ([string]$template.Name)
        Write-Host ".\QuestODM-Client-Runner-v2.ps1 $collectionText $templateText R"
    }
    return $templates
}

function Get-SelectedTemplate {
    param([AllowEmptyString()][string]$NameOrId)

    $templates = @(Get-AvailableTemplates)
    if ($templates.Count -eq 0) {
        throw "No saved $($script:Config.MigrationAsset) templates were found in the Accounts workload."
    }

    if ([string]::IsNullOrWhiteSpace($NameOrId)) {
        return Read-NumberedChoice -Title "Available $($script:Config.MigrationAsset) templates" -Items $templates -Display {
            param($item) "$($item.Name) | ID: $($item.Id)"
        }
    }

    $matches = @($templates | Where-Object {
        [string]$_.Name -eq $NameOrId -or [string]$_.Id -eq $NameOrId
    })
    if ($matches.Count -eq 0) {
        $matches = @($templates | Where-Object {
            ([string]$_.Name).IndexOf($NameOrId, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
        })
    }

    if ($matches.Count -ne 1) {
        $null = Show-AvailableTemplates -CollectionForCommand $CollectionName
        throw "Template '$NameOrId' matched $($matches.Count) templates. Use the full template name, its ID, or run with -ListTemplates."
    }
    return $matches[0]
}

function Write-ReusableMigrationCommand {
    param(
        [Parameter(Mandatory)][object]$Template,
        [Parameter(Mandatory)][ValidateSet('RunNow', 'CreateOnly', 'Schedule')][string]$Mode
    )

    $code = @{ RunNow = 'R'; CreateOnly = 'C'; Schedule = 'S' }[$Mode]
    $collectionText = ConvertTo-CommandArgument -Value $CollectionName
    $templateText = ConvertTo-CommandArgument -Value ([string]$Template.Name)
    Write-Host "Reusable command: .\QuestODM-Client-Runner-v2.ps1 $collectionText $templateText $code" -ForegroundColor DarkGray
}

function Assert-UsersAreMatched {
    param([Parameter(Mandatory)][object[]]$Users)

    $unmatched = @($Users | Where-Object {
        [string]::IsNullOrWhiteSpace([string](Get-PropertyValue $_ 'TargetUserPrincipalName'))
    })
    if ($unmatched.Count -gt 0) {
        throw "$($unmatched.Count) user(s) have no target UPN. Run the M operation first. No migration task was created."
    }
}

function New-MigrationTask {
    param(
        [Parameter(Mandatory)][object]$Template,
        [AllowNull()][Nullable[datetime]]$ScheduleUtc
    )

    $name = '{0}-{1}-{2}' -f $Template.Name, $CollectionName, (Get-Date -Format 'yyyyMMdd-HHmmss')
    $parameters = @{
        Name       = $name
        TemplateId = [string]$Template.Id
    }
    if ($null -ne $ScheduleUtc) {
        $parameters.ScheduledStartTime = [datetime]$ScheduleUtc
    }

    if ([string]$script:Config.MigrationAsset -eq 'Mailbox') {
        return New-OdmMailMigrationTask @parameters
    }
    return New-OdmOneDriveMigrationTask @parameters
}

function Wait-QuestTask {
    param(
        [Parameter(Mandatory)][object]$Task,
        [Parameter(Mandatory)][object]$Collection,
        [switch]$InProgressAlreadySent
    )

    $deadline = (Get-Date).AddHours([double]$script:Config.MaxMonitorHours)
    $lastStatus = ''
    $sentInProgress = [bool]$InProgressAlreadySent

    while ($true) {
        $fresh = @(Get-OdmTask -TaskId ([string]$Task.Id)) | Select-Object -First 1
        if ($null -eq $fresh) { throw "Task '$($Task.Id)' was not found while monitoring." }
        $status = [string](Get-PropertyValue $fresh 'Status')
        if ($status -ne $lastStatus) {
            Write-Host ("{0:yyyy-MM-dd HH:mm:ss} | {1} | {2}%" -f (Get-Date), $status, $fresh.Progress) -ForegroundColor Cyan
            $lastStatus = $status
        }

        if (-not $sentInProgress -and $status -in $script:InProgressTaskStatuses) {
            Send-TaskStatus -Stage 'In progress' -Task $fresh -Collection $Collection -Detail 'Quest is processing the task.'
            $sentInProgress = $true
        }

        if ($status -notin $script:ActiveTaskStatuses) {
            Send-TaskStatus -Stage 'Finished' -Task $fresh -Collection $Collection -Detail "Final Quest status: $status. Last result: $($fresh.LastResult)"
            return $fresh
        }
        if ((Get-Date) -ge $deadline) {
            throw "Monitoring exceeded $($script:Config.MaxMonitorHours) hours. Run MON later with task ID '$($Task.Id)'."
        }
        Start-Sleep -Seconds ([int]$script:Config.PollSeconds)
    }
}

function Invoke-MigrationOperation {
    param(
        [Parameter(Mandatory)][object]$Collection,
        [Parameter(Mandatory)][object]$Template,
        [Parameter(Mandatory)][ValidateSet('RunNow', 'CreateOnly', 'Schedule')][string]$Mode
    )

    $users = Get-CollectionUsers -Collection $Collection
    Assert-UsersAreMatched -Users $users
    Write-ReusableMigrationCommand -Template $Template -Mode $Mode
    $scheduleUtc = if ($Mode -eq 'Schedule') { Read-ScheduledTimeIst } else { $null }

    $null = New-RunFolder -Collection ([string]$Collection.Name) -Action $Mode
    $null = Export-CollectionSnapshot -Collection $Collection -Phase Pre

    $task = New-MigrationTask -Template $Template -ScheduleUtc $scheduleUtc
    Add-OdmObject -Objects $users -To $task
    Add-OdmTask -Tasks $task -To $Collection
    Send-TaskStatus -Stage 'Created' -Task $task -Collection $Collection -Detail "Mode: $Mode. Asset: $($script:Config.MigrationAsset)."

    if ($Mode -eq 'CreateOnly') {
        $null = Export-CollectionSnapshot -Collection $Collection -Phase 'Post-Created'
        Export-TaskDetails -Task $task
        Write-Host "Created for Run Later: $($task.Name) [$($task.Id)]" -ForegroundColor Green
        Write-Host 'Use MON after the task is started to create the final post report.' -ForegroundColor Green
        return
    }

    $inProgressSent = $false
    if ($Mode -eq 'RunNow') {
        Start-OdmTask -Tasks $task
        Send-TaskStatus -Stage 'In progress' -Task $task -Collection $Collection -Detail 'The task was submitted to Quest for immediate execution.'
        $inProgressSent = $true
    }
    else {
        $zone = Get-ConfiguredTimeZone
        $ist = [TimeZoneInfo]::ConvertTimeFromUtc([datetime]$scheduleUtc, $zone)
        Write-Host "Scheduled for $($ist.ToString('yyyy-MM-dd HH:mm')) IST." -ForegroundColor Green
        if (-not [bool]$script:Config.MonitorScheduledTasks) {
            $null = Export-CollectionSnapshot -Collection $Collection -Phase 'Post-Created'
            Export-TaskDetails -Task $task
            Write-Host "Scheduled: $($task.Name) [$($task.Id)]. Use MON later for the final post report." -ForegroundColor Green
            return
        }
    }

    try {
        $task = Wait-QuestTask -Task $task -Collection $Collection -InProgressAlreadySent:$inProgressSent
    }
    finally {
        $null = Export-CollectionSnapshot -Collection $Collection -Phase Post
        Export-TaskDetails -Task $task
    }
    Write-Host "Finished: $($task.Name) [$($task.Id)] | $($task.Status)" -ForegroundColor Green
}

function Test-MappingFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object[]]$Users
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Mapping CSV was not found: $Path"
    }
    $item = Get-Item -LiteralPath $Path
    if ($item.Length -gt 15MB) { throw 'Mapping CSV exceeds the Quest 15 MB limit.' }
    if ($item.Name -match '[^\x00-\x7F]') { throw 'Mapping CSV filename must contain ASCII characters only.' }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { throw 'Mapping CSV has no data rows.' }
    $headers = @($rows[0].PSObject.Properties.Name)
    $hasUpn = $headers -contains 'SourceUPN'
    $hasId = $headers -contains 'SourceObjectId'
    if (-not $hasUpn -and -not $hasId) {
        throw 'Mapping CSV requires SourceUPN or SourceObjectId.'
    }
    if (($headers -notcontains 'TargetUPN') -and ($headers -notcontains 'TargetObjectId')) {
        throw 'Mapping CSV requires TargetUPN or TargetObjectId.'
    }

    $allowedUpn = @{}
    $allowedId = @{}
    foreach ($user in $Users) {
        $allowedUpn[[string](Get-PropertyValue $user 'SourceUserPrincipalName')] = $true
        $allowedId[[string](Get-PropertyValue $user 'Id')] = $true
    }

    $outside = @($rows | Where-Object {
        $sourceUpn = if ($hasUpn) { [string]$_.SourceUPN } else { '' }
        $sourceId = if ($hasId) { [string]$_.SourceObjectId } else { '' }
        (-not [string]::IsNullOrWhiteSpace($sourceUpn) -and -not $allowedUpn.ContainsKey($sourceUpn)) -or
        (-not [string]::IsNullOrWhiteSpace($sourceId) -and -not $allowedId.ContainsKey($sourceId)) -or
        ([string]::IsNullOrWhiteSpace($sourceUpn) -and [string]::IsNullOrWhiteSpace($sourceId))
    })
    if ($outside.Count -gt 0) {
        throw "$($outside.Count) mapping row(s) are blank or outside collection '$CollectionName'."
    }
}

function New-MappingTask {
    param([Parameter(Mandatory)][object[]]$Users)

    $name = 'Accounts-Mapping-{0}-{1}' -f $CollectionName, (Get-Date -Format 'yyyyMMdd-HHmmss')
    $email = [string]$script:Config.NotificationRecipientEmail
    $common = @{
        Name                       = $name
        NotifyOnTaskCompletion     = -not [string]::IsNullOrWhiteSpace($email)
        NotifyOnlyOnTaskFailure    = [bool]$script:Config.NotifyOnlyOnFailure
        NotificationRecipientEmail = $email
    }

    if ([string]$script:Config.MappingMode -eq 'File') {
        $path = if (-not [string]::IsNullOrWhiteSpace($InputCsvPath)) {
            $InputCsvPath
        }
        else {
            [string]$script:Config.MappingFilePath
        }
        $path = [Environment]::ExpandEnvironmentVariables($path)
        Test-MappingFile -Path $path -Users $Users
        $common.MatchingAction = [string]$script:Config.MappingFileAction
        $common.FilePath = $path
        return New-OdmMappingFileTask @common
    }

    $common.MatchingAction = 'Matching'
    $common.SourceAttribute = [string]$script:Config.MatchingSourceAttribute
    $common.TargetAttribute = [string]$script:Config.MatchingTargetAttribute
    $common.MatchingRelation = [string]$script:Config.MatchingRelation
    $common.AssignOdmLicense = [bool]$script:Config.AssignOdmLicense
    return New-OdmMatchingTask @common
}

function Invoke-MappingOperation {
    param([Parameter(Mandatory)][object]$Collection)

    $users = Get-CollectionUsers -Collection $Collection
    $null = New-RunFolder -Collection ([string]$Collection.Name) -Action 'Mapping'
    $null = Export-CollectionSnapshot -Collection $Collection -Phase Pre

    $task = New-MappingTask -Users $users
    if ([string]$script:Config.MappingMode -eq 'Attribute') {
        Add-OdmObject -Objects $users -To $task
    }
    Add-OdmTask -Tasks $task -To $Collection
    Send-TaskStatus -Stage 'Created' -Task $task -Collection $Collection -Detail "Mapping mode: $($script:Config.MappingMode)."
    Start-OdmTask -Tasks $task
    Send-TaskStatus -Stage 'In progress' -Task $task -Collection $Collection -Detail 'The user mapping task was submitted to Quest.'

    try {
        $task = Wait-QuestTask -Task $task -Collection $Collection -InProgressAlreadySent
    }
    finally {
        $null = Export-CollectionSnapshot -Collection $Collection -Phase Post
        Export-TaskDetails -Task $task
    }
    Write-Host "Mapping finished: $($task.Name) [$($task.Id)] | $($task.Status)" -ForegroundColor Green
}

function Get-CollectionTask {
    param([Parameter(Mandatory)][object]$Collection)

    if (-not [string]::IsNullOrWhiteSpace($TaskId)) {
        $task = @(Get-OdmTask -TaskId $TaskId) | Select-Object -First 1
        if ($null -eq $task) { throw "Task '$TaskId' was not found." }
        return $task
    }

    $task = @(Get-OdmTask -CollectionId ([string]$Collection.Id) -All | Sort-Object Created -Descending) |
        Select-Object -First 1
    if ($null -eq $task) { throw "No task was found for collection '$($Collection.Name)'." }
    return $task
}

function Invoke-MonitorOperation {
    param([Parameter(Mandatory)][object]$Collection)

    $task = Get-CollectionTask -Collection $Collection
    $null = New-RunFolder -Collection ([string]$Collection.Name) -Action 'Monitor'
    $null = Export-CollectionSnapshot -Collection $Collection -Phase Pre
    try {
        $task = Wait-QuestTask -Task $task -Collection $Collection
    }
    finally {
        $null = Export-CollectionSnapshot -Collection $Collection -Phase Post
        Export-TaskDetails -Task $task
    }
    Write-Host "Monitoring finished: $($task.Name) [$($task.Id)] | $($task.Status)" -ForegroundColor Green
}

function Invoke-ReportOperation {
    param([Parameter(Mandatory)][object]$Collection)

    $null = New-RunFolder -Collection ([string]$Collection.Name) -Action 'Report'
    $null = Export-CollectionSnapshot -Collection $Collection -Phase Current
    $tasks = @(Get-OdmTask -CollectionId ([string]$Collection.Id) -All | Sort-Object Created -Descending)
    if (-not [string]::IsNullOrWhiteSpace($TaskId)) {
        $tasks = @($tasks | Where-Object { [string]$_.Id -eq $TaskId })
    }
    if ($tasks.Count -gt 0) {
        Export-TaskDetails -Task $tasks[0]
    }
    Write-Host "Report folder: $script:RunFolder" -ForegroundColor Green
}

function Get-CsvSourceValue {
    param([Parameter(Mandatory)][object]$Row)

    foreach ($name in 'UserPrincipalName', 'SourceUPN', 'S-UserPrincipalName', 'Email', 'SourceEmail', 'ObjectId') {
        if ($Row.PSObject.Properties.Name -contains $name -and
            -not [string]::IsNullOrWhiteSpace([string]$Row.$name)) {
            return [PSCustomObject]@{ Header = $name; Value = ([string]$Row.$name).Trim() }
        }
    }
    return $null
}

function Invoke-NewCollectionOperation {
    $path = if (-not [string]::IsNullOrWhiteSpace($InputCsvPath)) {
        $InputCsvPath
    }
    else {
        [string]$script:Config.CollectionInputCsvPath
    }
    $path = [Environment]::ExpandEnvironmentVariables($path)
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Collection input CSV was not found: $path"
    }

    $limit = if ($script:CollectionLimitWasSpecified) {
        $script:RequestedCollectionLimit
    }
    else {
        [int]$script:Config.CollectionLimit
    }
    if ($limit -lt 1) { throw 'Collection limit must be at least 1.' }

    $rows = @(Import-Csv -LiteralPath $path)
    if ($rows.Count -eq 0) { throw 'Collection input CSV has no data rows.' }
    $allObjects = @(Get-OdmObject -All -IncludeAllProperties -IncludeCollections)
    $allUsers = @($allObjects | Where-Object { [string]$_.Type -eq 'User' })

    $byId = @{}
    $byUpn = @{}
    $byEmail = @{}
    foreach ($user in $allUsers) {
        $id = [string](Get-PropertyValue $user 'Id')
        $upn = [string](Get-PropertyValue $user 'SourceUserPrincipalName')
        $email = [string](Get-PropertyValue $user 'SourceEmail')
        if (-not [string]::IsNullOrWhiteSpace($id)) { $byId[$id] = $user }
        if (-not [string]::IsNullOrWhiteSpace($upn)) { $byUpn[$upn] = $user }
        if (-not [string]::IsNullOrWhiteSpace($email)) { $byEmail[$email] = $user }
    }

    $plan = [System.Collections.ArrayList]::new()
    $selected = @{}
    foreach ($row in $rows) {
        $source = Get-CsvSourceValue -Row $row
        $user = $null
        if ($null -ne $source) {
            if ($source.Header -eq 'ObjectId') { $user = $byId[$source.Value] }
            elseif ($source.Header -in @('Email', 'SourceEmail')) { $user = $byEmail[$source.Value] }
            else { $user = $byUpn[$source.Value] }
        }
        $found = $null -ne $user
        $null = $plan.Add([PSCustomObject][ordered]@{
            InputHeader = if ($null -ne $source) { $source.Header } else { '' }
            InputValue  = if ($null -ne $source) { $source.Value } else { '' }
            FoundInODM  = $found
            ObjectId    = if ($found) { Get-PropertyValue $user 'Id' } else { '' }
            DisplayName = if ($found) { Get-PropertyValue $user 'DisplayName' } else { '' }
            SourceUPN   = if ($found) { Get-PropertyValue $user 'SourceUserPrincipalName' } else { '' }
        })
        if ($found) { $selected[[string](Get-PropertyValue $user 'Id')] = $user }
    }

    $null = New-RunFolder -Collection $CollectionName -Action 'NewCollection'
    $plan | Export-Csv -LiteralPath (Join-Path $script:RunFolder 'Collection-Plan-Pre.csv') -NoTypeInformation -Encoding UTF8
    $missing = @($plan | Where-Object { -not $_.FoundInODM })
    if ($missing.Count -gt 0) {
        throw "$($missing.Count) CSV row(s) were blank or not discovered in the Accounts workload. See Collection-Plan-Pre.csv. No collection was created."
    }

    $users = @($selected.Values | Sort-Object SourceUserPrincipalName)
    if ($users.Count -eq 0) { throw 'No unique user accounts were selected.' }
    $partCount = [int][math]::Ceiling($users.Count / [double]$limit)
    $names = @()
    for ($index = 0; $index -lt $partCount; $index++) {
        if ($partCount -eq 1) { $names += $CollectionName }
        else { $names += ('{0}-{1:d3}' -f $CollectionName, ($index + 1)) }
    }

    $existing = @(Get-OdmCollection -ResultSize 10000)
    $conflicts = @($names | Where-Object { $name = $_; @($existing | Where-Object { [string]$_.Name -eq $name }).Count -gt 0 })
    if ($conflicts.Count -gt 0) {
        throw "Collection name(s) already exist: $($conflicts -join ', '). No collection was created."
    }

    $result = [System.Collections.ArrayList]::new()
    for ($index = 0; $index -lt $partCount; $index++) {
        $start = $index * $limit
        $end = [math]::Min($start + $limit - 1, $users.Count - 1)
        $partUsers = @($users[$start..$end])
        $collection = New-OdmCollection -Name $names[$index]
        Add-OdmObject -Objects $partUsers -To $collection
        $null = $result.Add([PSCustomObject][ordered]@{
            CollectionName = $collection.Name
            CollectionId   = $collection.Id
            UserCount      = $partUsers.Count
            Limit          = $limit
            Created        = Get-Date
        })
        Write-Host "Created '$($collection.Name)' with $($partUsers.Count) user(s)." -ForegroundColor Green
    }
    $result | Export-Csv -LiteralPath (Join-Path $script:RunFolder 'Collection-Result-Post.csv') -NoTypeInformation -Encoding UTF8
    Write-Host "Collection reports: $script:RunFolder" -ForegroundColor Green
}

# Main
# Preserve the earlier two-value form: "Collection" R. It now opens the
# template list for migration operations instead of silently using a default.
if ([string]::IsNullOrWhiteSpace($Operation) -and (Test-OperationCode -Value $TemplateName)) {
    $Operation = $TemplateName
    $TemplateName = ''
}

$script:CollectionLimitWasSpecified = $PSBoundParameters.ContainsKey('CollectionLimit')
$script:RequestedCollectionLimit = $CollectionLimit
$script:Config = Import-ClientConfig

if (-not $ListTemplates -and [string]::IsNullOrWhiteSpace($Operation)) {
    $Operation = Read-OperationFromMenu
}
$resolvedOperation = if (-not $ListTemplates) {
    Resolve-OperationCode -Value $Operation
}
else {
    $null
}

Connect-Quest

if ($ListTemplates) {
    $null = Show-AvailableTemplates -CollectionForCommand $CollectionName
    return
}

if ($resolvedOperation -eq 'NewCollection') {
    if ([string]::IsNullOrWhiteSpace($CollectionName)) {
        $CollectionName = (Read-Host 'New collection base name').Trim()
    }
    if ([string]::IsNullOrWhiteSpace($CollectionName)) {
        throw 'A new collection base name is required.'
    }
    Invoke-NewCollectionOperation
    return
}

if ([string]::IsNullOrWhiteSpace($CollectionName)) {
    $collection = Select-AccountsCollectionFromMenu
}
else {
    $collection = Find-AccountsCollection -Name $CollectionName
}
$CollectionName = [string]$collection.Name

switch ($resolvedOperation) {
    'RunNow'    {
        $template = Get-SelectedTemplate -NameOrId $TemplateName
        Invoke-MigrationOperation -Collection $collection -Template $template -Mode RunNow
    }
    'CreateOnly'{
        $template = Get-SelectedTemplate -NameOrId $TemplateName
        Invoke-MigrationOperation -Collection $collection -Template $template -Mode CreateOnly
    }
    'Schedule'  {
        $template = Get-SelectedTemplate -NameOrId $TemplateName
        Invoke-MigrationOperation -Collection $collection -Template $template -Mode Schedule
    }
    'Mapping'   { Invoke-MappingOperation -Collection $collection }
    'Monitor'   { Invoke-MonitorOperation -Collection $collection }
    'Report'    { Invoke-ReportOperation -Collection $collection }
}
