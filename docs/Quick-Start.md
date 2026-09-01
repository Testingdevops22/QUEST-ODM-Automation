# Quest ODM Client Runner v2 — Quick Start

## What changed

Version 2 is a separate, command-first script. Version 1 is unchanged.

Only **Accounts users** are used. Other workloads and non-user objects are not offered to the operator. **Mailbox** or **OneDrive** is predefined in configuration.

## One-time setup

1. Install/update Quest's module in Windows PowerShell 5.1:

```powershell
Install-Module OdmApi -Scope CurrentUser
Update-Module OdmApi
```

2. Place the runner and working config together:

```text
QuestODM-Client-Runner-v2.ps1
QuestODM-Client-Config-v2.txt
```

3. Edit the config once:

- `OrganizationId`: Quest organization GUID
- `ProjectId`: parent migration project ID
- `MigrationAsset`: `Mailbox` or `OneDrive`
- `NotificationRecipientEmail`: semicolon-separated recipients if needed
- `OutputDirectory`: report folder

If the project has more than one Accounts workload, set `AccountsWorkloadId`; otherwise leave it blank.

4. Never add passwords, secrets, tokens, or credentials to the config. Version 2 uses Windows sign-in when available, otherwise one Microsoft MFA sign-in, then reuses that Quest session for the command.

## Daily use

Normal command:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" "Preload" R
```

Argument order:

1. Accounts collection name
2. Saved Quest template name (for `R`, `C`, or `S`)
3. Operation

Run now:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" "Preload" R
```

Finalization:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" "Finalization" R
```

Create only / Run Later:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" "Preload" C
```

Schedule using the IST picker:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" "Preload" S
```

## Find saved template names

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" -ListTemplates
```

The script lists saved Mailbox or OneDrive templates and prints reusable one-line commands. A full template ID can also be supplied instead of the template name.

## Automatic menus

If a value is missing:

- Missing operation → numbered operation menu
- Missing collection → numbered Accounts collection menu
- Missing template for `R`, `C`, or `S` → numbered saved-template menu

Examples:

```powershell
.\QuestODM-Client-Runner-v2.ps1
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" R
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" "Preload"
.\QuestODM-Client-Runner-v2.ps1 -TemplateName "Preload" -Operation R
```

## Operations that do not require a template

Map users:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" M
```

Monitor latest task:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" MON
```

Monitor a specific task:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" MON -TaskId "task-id"
```

Export current report:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" REP
```

Long operation names also work: `RUN`, `CREATE`, `SCHEDULE`, `MAPPING`, `MONITOR`, `REPORT`, and `NEW`.

## Reports and email

`R`, `S`, and `M` create a pre-report before task creation and a post-report after completion. Task summary, task events, and status history are stored in the same timestamped folder.

Status emails are sent for:

- Created
- In progress
- Finished

These use the already signed-in classic Outlook desktop profile. No mailbox password is stored.

Disable these messages with:

```json
"StatusEmailMode": "None"
```

For scheduled tasks, keep PowerShell open if `MonitorScheduledTasks` is enabled and you want continuous monitoring. Closing the window does not remove the Quest schedule; later run `MON` to monitor it and generate the final post-report.

For `C` (Create task only), Version 2 creates a pre-report and Post-Created snapshot. After the task is manually started in Quest, run `MON` to generate the final post-report and completion alert.

## User mapping

Attribute matching:

```json
"MappingMode": "Attribute",
"MatchingSourceAttribute": "Mail",
"MatchingTargetAttribute": "Mail",
"MatchingRelation": "Priority"
```

Mapping file:

```json
"MappingMode": "File",
"MappingFilePath": "C:\\QuestODM\\Input\\UserMapping.csv",
"MappingFileAction": "MatchingFromFile"
```

Override for one command:

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" M -InputCsvPath "C:\QuestODM\Input\UserMapping.csv"
```

Mapping files must use either:

- `SourceUPN` / `TargetUPN`
- `SourceObjectId` / `TargetObjectId`

Every source user must already be in the named collection. The script rejects files larger than 15 MB and non-ASCII filenames.

## Create collections with a limit

Accepted source columns:

```text
UserPrincipalName
SourceUPN
S-UserPrincipalName
Email
SourceEmail
ObjectId
```

Example CSV:

```text
UserPrincipalName
user1@source.example
user2@source.example
```

Create collection(s):

```powershell
.\QuestODM-Client-Runner-v2.ps1 "Wave 01" NEW -InputCsvPath "C:\QuestODM\Input\Wave01.csv" -CollectionLimit 200
```

When the CSV exceeds the limit, the script creates names such as:

```text
Wave 01-001
Wave 01-002
```

It validates the CSV rows and future collection names before creating anything. Only discovered Accounts users are added.

## Safety notes

- `R`, `C`, and `S` stop if any user has no Target UPN; run mapping first.
- The selected collection must be an Accounts collection.
- Version 2 never stores a password, secret, token, or credential.
- The saved Quest template remains the source of migration settings.
- Quest completion notification settings stored in that template are retained.
