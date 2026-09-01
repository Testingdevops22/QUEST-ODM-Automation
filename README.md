# QUEST ODM Automation

PowerShell automation for Quest On Demand Migration (ODM) focused on Microsoft 365 tenant-to-tenant migration workflows for **Exchange Online (EXO)** and **OneDrive for Business (ODB)**.

## What this project does

The Version 2 client runner is command-first and works only with **Accounts users**. The migration workload is predefined in configuration as either **Mailbox** or **OneDrive**.

Supported operator actions include:

- Create Accounts collections from CSV input
- Map/match source and target users
- List saved Quest migration templates
- Create migration tasks without starting them
- Start migration tasks immediately
- Schedule migration tasks
- Monitor the latest or a specific task
- Export migration reports
- Generate pre/post task snapshots
- Send status notifications through the signed-in Outlook desktop profile

## Repository structure

```text
QUEST-ODM-Automation/
├── README.md
├── .gitignore
├── src/
│   └── QuestODM-Client-Runner-v2.ps1
├── config/
│   └── QuestODM-Client-Config-v2.example.txt
└── docs/
    └── Quick-Start.md
```

## Requirements

- Windows PowerShell 5.1
- Quest `OdmApi` PowerShell module
- Access to the required Quest On Demand Migration organization/project
- Saved migration templates in Quest ODM
- Classic Outlook desktop profile if Outlook-based status email is enabled

Install/update the Quest module:

```powershell
Install-Module OdmApi -Scope CurrentUser
Update-Module OdmApi
```

## Initial configuration

Copy:

```text
config/QuestODM-Client-Config-v2.example.txt
```

to a working file such as:

```text
QuestODM-Client-Config-v2.txt
```

Then update at minimum:

- `OrganizationId`
- `ProjectId`
- `UserPrincipalName`
- `MigrationAsset` (`Mailbox` or `OneDrive`)
- `NotificationRecipientEmail`
- `OutputDirectory`

If the project contains multiple Accounts workloads, set `AccountsWorkloadId`. Otherwise leave it blank.

> Do not store passwords, secrets, tokens, or credentials in the configuration file.

## Basic usage

General command format:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "<CollectionName>" "<TemplateName>" <Operation>
```

Example — run now:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" "Preload" R
```

Example — finalization:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" "Finalization" R
```

Example — create task only:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" "Preload" C
```

Example — schedule task:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" "Preload" S
```

## Operations

| Operation | Purpose |
|---|---|
| `R` / `RUN` | Create and start a task now |
| `C` / `CREATE` | Create task only |
| `S` / `SCHEDULE` | Schedule a task |
| `M` / `MAPPING` | Match/map users |
| `MON` / `MONITOR` | Monitor task |
| `REP` / `REPORT` | Export current report |
| `NEW` | Create collection(s) from CSV |

## List templates

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" -ListTemplates
```

The runner lists saved templates for the configured Mailbox or OneDrive workload and prints reusable one-line commands.

## User mapping

Attribute-based matching example:

```json
"MappingMode": "Attribute",
"MatchingSourceAttribute": "Mail",
"MatchingTargetAttribute": "Mail",
"MatchingRelation": "Priority"
```

CSV-based matching example:

```json
"MappingMode": "File",
"MappingFilePath": "C:\\QuestODM\\Input\\UserMapping.csv",
"MappingFileAction": "MatchingFromFile"
```

Or override the file for a single command:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" M -InputCsvPath "C:\QuestODM\Input\UserMapping.csv"
```

Mapping files must use either:

- `SourceUPN` / `TargetUPN`
- `SourceObjectId` / `TargetObjectId`

Every source user must already belong to the selected Accounts collection.

## Collection creation

Accepted source columns include:

- `UserPrincipalName`
- `SourceUPN`
- `S-UserPrincipalName`
- `Email`
- `SourceEmail`
- `ObjectId`

Example:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" NEW -InputCsvPath "C:\QuestODM\Input\Wave01.csv" -CollectionLimit 200
```

If the CSV exceeds the configured collection limit, the runner creates numbered collections such as:

```text
Wave 01-001
Wave 01-002
```

## Reporting and monitoring

For `R`, `S`, and `M`, the runner creates a pre-report before task creation and a post-report after completion. Task summary, task events, and status history are written to the same timestamped report folder.

Status email events include:

- Created
- In progress
- Finished

Set:

```json
"StatusEmailMode": "None"
```

to disable Outlook notifications.

For scheduled tasks, keep the PowerShell process open when `MonitorScheduledTasks` is enabled if you want continuous monitoring. Closing PowerShell does not remove the Quest schedule; you can later run `MON` to resume monitoring and generate the final post-report.

## Safety behavior

The runner includes important execution guards:

- `R`, `C`, and `S` stop when any collection user has no Target UPN
- The selected collection must be an Accounts collection
- Passwords, secrets, tokens, and credentials are not stored by Version 2
- The saved Quest template remains the source of migration task settings
- Existing Quest completion notification settings in the selected template are retained

## EXO vs ODB

Set the workload through:

```json
"MigrationAsset": "Mailbox"
```

for Exchange Online, or:

```json
"MigrationAsset": "OneDrive"
```

for OneDrive migrations.

Use separate working configuration files if you regularly run both workloads, for example:

```text
QuestODM-EXO-Config.txt
QuestODM-ODB-Config.txt
```

## Security

This public repository intentionally contains only a sanitized example configuration. Do not commit:

- Tenant-specific secrets
- Access tokens
- Passwords
- Client credentials
- Real organization/project identifiers unless explicitly intended
- User mapping CSV files containing client data
- Generated migration reports

## Disclaimer

This project is an independent automation utility built around Quest On Demand Migration PowerShell capabilities. It is not an official Quest Software or Microsoft product. Test all changes in a non-production migration project before production use.
