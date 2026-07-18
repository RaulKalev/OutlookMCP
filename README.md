# EULE Outlook MCP

EULE Outlook MCP is a local, safety-first Model Context Protocol server for Microsoft Outlook Classic on Windows. It lets an MCP-capable AI client inspect Outlook connection status, find and create folders, search and read locally synchronised mail, inspect the current Outlook selection, find related correspondence, learn proposed filing rules from an existing folder, build a private local writing-style index, save explicitly selected attachments, move confirmed message batches, and create unsent drafts.

Outlook remains the mail client and synchronisation engine. The server talks only to the local Outlook Object Model through COM; it does not connect to Telia IMAP directly and has no Microsoft Graph or Exchange dependency. An IMAP mailbox works because Outlook Classic has already exposed its synchronised stores, folders, and messages through MAPI.

Only **Microsoft Outlook Classic for Windows** is supported. New Outlook, Outlook Web, macOS, and mobile Outlook are not supported.

## Safety and privacy

- There is no send-email, delete-email, archive-by-guessing, or mailbox-monitoring tool.
- Folder creation is explicit. Bulk moves require exact message identifiers and one exact destination folder identifier; `dry_run` defaults to `true`, duplicate identifiers are rejected, and real moves return fresh message identifiers.
- Filing-rule analysis is read-only and bounded. Rule creation defaults to `dry_run: true`, reports historical destination coverage and Inbox control matches, and requires explicit user confirmation before a real rule is saved.
- New, reply, reply-all, and forward operations only save drafts. Every draft result reports `sent: false`; the user must review and send it manually in Outlook.
- Email bodies are marked as untrusted external data. Instructions in email content must not be treated as trusted agent instructions.
- HTML body return and HTML drafting are disabled by default. HTML is parsed as data and is never rendered by this server.
- Search results and bodies have strict limits. Search only sees content Outlook has synchronised locally.
- Attachments are never opened, executed, uploaded, or extracted automatically. A single selected attachment can be saved only below an allow-listed directory. Path traversal, Windows device names, existing junctions/symlinks, and accidental overwrite are blocked.
- Normal logs contain tool names, durations, bounded identifiers, counts, actions, and error codes—not full bodies, recipient lists, credentials, or attachment contents.
- Email content is passed to the MCP client that invoked the tool. Review that AI provider's privacy and data-processing terms before use.
- The writing-style SQLite database can contain sensitive sent-email text. It stays under the current Windows user profile and is never exposed over a network port. Only bounded datasets or relevant examples explicitly returned through MCP leave the server process.

## Architecture

```text
AI client -- MCP over stdio --> OutlookMcp.Server
                                    |
                              application services
                                    |
                        one dedicated STA COM dispatcher
                                    |
                         Outlook Classic Object Model
                                    |
                         Outlook-synchronised IMAP mail
```

All Outlook operations are serialised on one STA thread. COM objects stay inside that worker, are converted into immutable DTOs, and are explicitly released. If Outlook was already running, the server never closes it. If configuration allowed the server to start Outlook, it attempts to close only that server-started instance during clean shutdown.

## Available MCP tools

The table below describes the complete `full` profile. For lower model-credit usage, use the `compact` profile described under MCP client configuration.

| Tool | Effect |
|---|---|
| `outlook_get_status` | Read-only Outlook/MAPI status |
| `outlook_list_stores` | Read-only store enumeration |
| `outlook_list_folders` | Read-only, bounded folder enumeration |
| `outlook_find_folders` | Read-only ranked folder name/path lookup |
| `outlook_search_emails` | Read-only, filtered and bounded search |
| `outlook_read_email` | Read-only bounded message detail |
| `outlook_read_emails_batch` | Read-only bounded multi-message detail with per-item errors |
| `outlook_read_thread` | Read-only bounded conversation assembly |
| `outlook_get_selected_email` | Read-only active Explorer/Inspector selection |
| `outlook_find_related_emails` | Read-only deterministic relevance search |
| `outlook_list_attachments` | Read-only attachment metadata |
| `outlook_save_attachment` | Saves one file to an approved local directory |
| `outlook_create_folder` | Creates one mail folder below an exact parent or store root |
| `outlook_move_emails` | Dry-run-first bulk move with fresh post-move identifiers |
| `outlook_analyze_folder_for_rules` | Read-only representative folder sample and recurring sender signals |
| `outlook_create_folder_rule` | Dry-run-first receive rule that moves future matching mail |
| `outlook_list_calendars` | Read-only calendar folder enumeration with default-calendar marking |
| `outlook_sync_calendar` | Dry-run-first one-way calendar sync into a dedicated target calendar |
| `outlook_create_draft` | Creates an unsent draft |
| `outlook_create_reply_draft` | Creates an unsent Reply/Reply All draft |
| `outlook_create_forward_draft` | Creates an unsent forward draft |
| `outlook_style_get_scan_status` | Local index progress and quality, without bodies |
| `outlook_style_scan_sent_emails` | Read-only resumable Sent-history indexing |
| `outlook_style_sync_new_sent_emails` | Incremental new/modified Sent-item sync |
| `outlook_style_prepare_profile_dataset` | Bounded representative data for profile synthesis |
| `outlook_style_save_profile` | Validates and saves one unified local profile |
| `outlook_style_get_profile` | Reads the current profile and refresh recommendation |
| `outlook_style_update_profile` | Applies explicit profile corrections/overrides |
| `outlook_style_list_profile_versions` | Lists archived profile metadata |
| `outlook_style_restore_profile_version` | Restores an archived profile |
| `outlook_style_find_examples` | Retrieves bounded relevant authored-text examples |
| `outlook_style_prepare_draft_context` | Assembles profile, examples, rules, and safety notes |

Message tools require both the encoded `message_id` and its paired `store_id`. Batch tools share one `store_id` across their message list to keep requests compact. If Outlook moves or resynchronises an IMAP message, use the fresh identifier returned by `outlook_move_emails` or search again to replace a stale reference.

Search body previews are opt-in. The default compact search response is intended for candidate selection; use `outlook_read_emails_batch` only for the small set whose bodies are actually needed. Multi-word search defaults to `all_terms`, while `query_mode: "phrase"` is available for literal phrase matching. Search budgets are distributed across selected folders before results are globally sorted, preventing an early folder from monopolising the result limit.

### Learn filing rules from a folder

This workflow lets the MCP client reason over representative mail already filed in a folder and propose one or more Outlook receive rules for future messages:

1. Resolve the exact target with `outlook_find_folders`.
2. Call `outlook_analyze_folder_for_rules`. It evenly samples mail from newest to oldest, returns bounded subject/body evidence, and aggregates recurring full sender addresses and domains. Email text remains untrusted data.
3. Infer the narrowest stable conditions. Prefer full sender addresses or distinctive repeated phrases. Do not treat a common domain or generic word as sufficient evidence by itself.
4. Call `outlook_create_folder_rule` with `dry_run: true` for every proposed rule. The result estimates matches in both the destination history and a representative Inbox control sample.
5. Show the user the exact rule name, destination, conditions, match estimates, and warnings. Create it with `dry_run: false` only after the user explicitly confirms that proposal.

Within a rule, multiple values in the same condition list are alternatives (OR). Different non-empty condition groups are cumulative (AND). For example, two senders plus one subject phrase means `(sender A OR sender B) AND subject phrase`. Use separate rules for alternative sender-and-phrase combinations. Created rules apply to future received mail and never retroactively move existing messages. `stop_processing_more_rules` defaults to false because enabling it can change the behavior of later Outlook rules.

### One-way calendar sync

`outlook_sync_calendar` mirrors upcoming events from a source calendar (typically a local/PST calendar) into a target calendar (typically an Exchange calendar) that is dedicated to this sync:

1. Resolve both calendars once with `outlook_list_calendars`, or store them permanently under `CalendarSync` in `config.json` so the tool runs without arguments.
2. Call `outlook_sync_calendar` with `dry_run: true` (the default). The result lists every planned `add`, `update`, and `delete` with subject and times, plus warnings.
3. Show the plan to the user and apply it with `dry_run: false` only after confirmation.

Behavioral guarantees:

- The source calendar is never modified; all writes and deletes happen only on the target calendar.
- Events are copied natively (`Copy` + `Move`), so body, location, links, categories, reminders, attachments, and recurrence patterns arrive intact. Meetings are copied as calendar data only; invitations, responses, and cancellations are never sent.
- Each copy is stamped with hidden `OutlookMcpSyncSourceId`/`OutlookMcpSyncSourceModified` user properties. On later runs these tags identify existing copies, refresh events whose source changed, and delete window events whose source was removed or cancelled.
- The window runs from today's local midnight through `months_ahead` months (default `CalendarSync.DefaultMonthsAhead`, 3). Target events entirely outside the window are left alone, so history is preserved.
- Recurring series are matched while active anywhere in the window and copied as complete series, including exceptions; an open-ended series therefore also places occurrences beyond the window boundary. This is reported as a warning.
- Because the target calendar is treated as sync-owned, in-window target events without a sync tag are also pruned; the dry run lists them first.
- If either calendar exceeds `CalendarSync.MaximumItemsScanned`, the sync refuses to run rather than compare an incomplete picture.

## Prerequisites

- Windows 10 or 11.
- Microsoft Outlook Classic installed with a working mail profile.
- For source builds: .NET 8 SDK.
- Outlook should be fully opened at least once, with the Telia IMAP account synchronised.

The release is self-contained and does not require a separately installed .NET runtime. Use the x64 build for normal 64-bit Office installations. An x86 release is also produced for legacy 32-bit environments. Outlook COM is out-of-process, but matching the installed Office bitness is the safest diagnostic choice.

## Build and test

```powershell
dotnet restore OutlookMcp.sln
dotnet build OutlookMcp.sln -c Release
dotnet test OutlookMcp.sln -c Release
```

Normal test runs execute pure unit tests and skip Outlook integration tests. See [Integration testing](#integration-testing) for explicit opt-in.

Create self-contained x64/x86 release directories and ZIP archives:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish.ps1
```

If Inno Setup 6 (`ISCC.exe`) is installed, the script also produces a per-user Windows installer. Otherwise the ZIP is a complete portable release.

## Install

### Quick install (recommended)

One script installs the server and registers it with every AI agent it detects on the machine — Claude Code, Claude Desktop, Codex, and Google Antigravity:

```powershell
# From a source checkout (builds with the .NET 8 SDK, installs to %LocalAppData%\Programs\EULE Outlook MCP):
powershell -ExecutionPolicy Bypass -File .\scripts\install-mcp.ps1

# From an extracted release ZIP or the installed folder (uses the executable next to the script):
powershell -ExecutionPolicy Bypass -File .\install-mcp.ps1
```

The installer offers the same registration as a post-install step, and the Start-menu shortcut **Register with AI Agents** reruns it at any time (for example after adding a new agent). Every configuration file is backed up next to the original (`*.backup-<timestamp>`) before it is changed, and re-running the script updates entries in place. Restart each agent afterwards; MCP configuration is read at startup.

Useful options:

| Option | Effect |
|---|---|
| `-Clients codex,antigravity` | Configure only the listed clients (`claude-code`, `claude-desktop`, `codex`, `antigravity`); also forces setup for clients that were not auto-detected |
| `-Clients all` | Force all supported clients, creating configuration files that do not exist yet |
| `-DryRun` | Show planned changes without writing anything |
| `-Remove` | Remove the `outlook` entry from the selected clients (also run automatically by the uninstaller) |
| `-ExePath <path>` | Register a specific executable (for example the x86 build) |
| `-Rebuild` | Rebuild from source even when an installed executable exists |
| `-Runtime win-x86` | Build the 32-bit executable when building from source |
| `-ToolProfile compact` | Choose `compact`, `mail`, `style`, or `full`; defaults to the lowest-credit `compact` profile |

The script only registers the server; the sections under [MCP client configuration](#mcp-client-configuration) document what it writes so the same setup can be done manually.

### Installer

1. Run `artifacts\installer\EULE-Outlook-MCP-Setup-1.2.0.exe`.
2. No administrator elevation is required; the default location is `%LocalAppData%\Programs\EULE Outlook MCP`.
3. The final page offers **Register with detected AI agents (Claude, Codex, Antigravity)**; leave it ticked to configure the clients immediately.
4. The installer creates Start-menu shortcuts for diagnostics, agent registration, and this README, and registers an uninstaller. Uninstalling also removes the server entry from client configurations.

### Portable ZIP

1. Extract `artifacts\EULE-Outlook-MCP-1.2.0-win-x64.zip` to a stable per-user folder.
2. Run `powershell -ExecutionPolicy Bypass -File .\install-mcp.ps1` from that folder, or configure clients manually.
3. Do not move the executable after configuring an MCP client unless you update the client command path (or rerun `install-mcp.ps1`).
4. Run the diagnostic command once before adding it to a client.

On first start the server creates:

- Configuration: `%AppData%\EULE\OutlookMcp\config.json`
- Logs: `%LocalAppData%\EULE\OutlookMcp\Logs`
- Writing-style index/profile: `%AppData%\EULE\OutlookMcp\WritingStyle`

## Diagnostics

The diagnostic mode does not read or print email bodies and does not create a draft:

```powershell
OutlookMcp.Server.exe --diagnose
OutlookMcp.Server.exe --version
OutlookMcp.Server.exe --print-config-path
OutlookMcp.Server.exe --print-log-path
```

`--diagnose` validates configuration, Outlook Classic COM activation, MAPI/profile availability, store enumeration, Drafts access, Explorer/Inspector access, and attachment-directory write access.

## MCP client configuration

`install-mcp.ps1` performs all of the registrations below automatically. This section documents the manual equivalent for each client and for MCP clients the script does not know about. Replace the executable path below with the installed path. Restart the MCP client after editing its configuration.

### Claude Code

```powershell
claude mcp add --scope user outlook -- "C:\Users\YOUR_NAME\AppData\Local\Programs\EULE Outlook MCP\OutlookMcp.Server.exe" --tool-profile compact
```

`--scope user` makes the server available in every project. Restart open Claude Code sessions and verify with `/mcp`.

### Codex app, CLI, or IDE extension

Add the server through **Settings > MCP servers > Add server**, choose **STDIO**, and select the executable. Alternatively, add this to `~/.codex/config.toml`:

```toml
[mcp_servers.outlook]
command = 'C:\Users\YOUR_NAME\AppData\Local\Programs\EULE Outlook MCP\OutlookMcp.Server.exe'
args = ["--tool-profile", "compact"]
startup_timeout_sec = 30
tool_timeout_sec = 60
enabled = true
```

Then restart Codex and use `/mcp` or `codex mcp list` to confirm the connection.

### Google Antigravity

Add the server through the agent panel's **... > Manage MCP Servers > View raw config**, or merge this into `%USERPROFILE%\.gemini\antigravity\mcp_config.json` (see [examples/antigravity-mcp-config.json](examples/antigravity-mcp-config.json)):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\Users\\YOUR_NAME\\AppData\\Local\\Programs\\EULE Outlook MCP\\OutlookMcp.Server.exe",
      "args": ["--tool-profile", "compact"]
    }
  }
}
```

Then refresh the MCP servers panel in Antigravity. If your Antigravity version shares a central config at `%USERPROFILE%\.gemini\config\mcp_config.json`, `install-mcp.ps1` updates that file as well when it exists.

### Claude Desktop and JSON-based MCP clients

Merge this into the client's MCP configuration (for Claude Desktop on Windows: `%APPDATA%\Claude\claude_desktop_config.json`; fully quit Claude Desktop from the tray before restarting):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\Users\\YOUR_NAME\\AppData\\Local\\Programs\\EULE Outlook MCP\\OutlookMcp.Server.exe",
      "args": ["--tool-profile", "compact"]
    }
  }
}
```

The server writes protocol messages only to stdout. Structured logs go to rolling local files, so they do not corrupt stdio transport.

### Tool profiles and model-credit usage

MCP clients commonly include every advertised tool schema in model context. Use `--tool-profile compact` for normal search, bounded batch reading, selected-mail access, related-mail lookup, attachment listing, calendar sync, and unsent draft creation. It advertises eleven smaller tools and returns reduced search/read DTOs with conservative defaults.

Other profiles are available when a task needs them:

| Profile | Advertised tools | Use for |
|---|---:|---|
| `compact` | 11 | Normal reading, drafting, and calendar sync with the lowest schema overhead |
| `mail` | 21 | Advanced search/read options, folder moves, filing-rule learning, calendar sync, and attachment saving |
| `style` | 11 | Writing-style index and profile maintenance only |
| `full` | 32 | Every capability; default when `--tool-profile` is omitted |

Keep `compact` configured most of the time and temporarily switch the argument to `mail`, `style`, or `full` only when that capability is needed. The profile changes only which MCP tools are advertised; server-side safety checks remain unchanged.

## Configuration

See [config.sample.json](src/OutlookMcp.Server/config.sample.json). Configuration is per user and validated on startup.

```json
{
  "Outlook": {
    "StartIfNotRunning": true,
    "AllowedStores": [],
    "BlockedFolderPaths": ["\\Personal"],
    "DefaultSearchFolders": ["Inbox", "Sent Items"],
    "DefaultSearchLimit": 25,
    "MaximumSearchLimit": 100,
    "DefaultBodyCharacterLimit": 50000,
    "OperationTimeoutSeconds": 30,
    "AllowHtmlBody": false,
    "AllowAttachmentSaving": true,
    "AllowedAttachmentDirectories": ["%USERPROFILE%\\Downloads\\Outlook AI"],
    "AllowSelectedEmailAccess": true,
    "MaximumRecursiveFolders": 1000,
    "MaximumBatchSize": 100
  },
  "CalendarSync": {
    "SourceCalendarFolderId": null,
    "SourceStoreId": null,
    "TargetCalendarFolderId": null,
    "TargetStoreId": null,
    "DefaultMonthsAhead": 3,
    "MaximumMonthsAhead": 24,
    "MaximumItemsScanned": 2500
  },
  "WritingStyle": {
    "Enabled": true,
    "DatabasePath": "%APPDATA%\\EULE\\OutlookMcp\\WritingStyle\\sent-email-index.db",
    "ProfilePath": "%APPDATA%\\EULE\\OutlookMcp\\WritingStyle\\writing-profile.json",
    "ProfileHistoryPath": "%APPDATA%\\EULE\\OutlookMcp\\WritingStyle\\writing-profile-history.json",
    "ScanAllSentFolders": true,
    "AllowedStores": [],
    "BatchSize": 100,
    "DelayBetweenBatchesMilliseconds": 100,
    "SaveCheckpointAfterEachBatch": true,
    "SyncBeforeDraftContext": true,
    "MaximumRetrievedExamples": 5,
    "MaximumCharactersPerExample": 3000,
    "MaximumDraftContextCharacters": 30000,
    "AllowFullAuthoredTextInResponses": false,
    "StoreCleanBody": true,
    "StoreQuotedText": true,
    "StoreSignatureText": true,
    "AdditionalSignatureMarkers": [],
    "AdditionalDisclaimerMarkers": [],
    "EnableFullTextSearch": true,
    "ProfileRefreshRecommendationThreshold": 100,
    "RecencyWeighting": {
      "RecentMonths": 12,
      "RecentWeight": 1.0,
      "MiddleYears": 3,
      "MiddleWeight": 0.8,
      "OlderWeight": 0.6
    }
  },
  "Logging": {
    "Level": "Information",
    "RetentionDays": 14,
    "IncludeTechnicalDetails": false
  }
}
```

An empty `AllowedStores` list exposes all profile stores except blocked folders. To restrict stores, enter exact `store_id` values from `outlook_list_stores` or exact display names. Environment variables in attachment directories are expanded. Existing filesystem links/junctions in attachment destination paths are rejected.

The `CalendarSync` identifiers are optional defaults for `outlook_sync_calendar`; fill them with values from `outlook_list_calendars` so a plain "sync my calendars" request needs no arguments. Explicit tool arguments always override these defaults.

## Local writing-style system

The writing-style feature processes every locally available item in every allowed Outlook Sent folder. It applies no date, message-length, recipient, project, language, or attachment filter. Every item produces an index row or a recorded failure/unsupported status. Reply history, forwarded content, signatures, and disclaimers are stored separately from likely user-authored text so they do not incorrectly define the user's style. Low-confidence and no-authored-text messages remain accounted for.

This is retrieval and structured profiling, not model training or fine-tuning. The server maintains one unified profile whose primary language is Estonian; conditional audience or language habits remain rules within that one profile. SQLite FTS5 supplies local full-text retrieval, then recipient, project, intent, confidence, recency, and duplicate-diversity signals rank a small example set for each draft.

### First use

1. Call `outlook_style_get_scan_status`.
2. Repeatedly call `outlook_style_scan_sent_emails` until `complete` is true. `maximum_batches` limits one call only; it never caps total mailbox coverage.
3. Call `outlook_style_prepare_profile_dataset`. The returned examples are explicitly marked `untrusted_data`.
4. Ask the AI agent to synthesise a version-1 structured profile from those statistics/examples, then call `outlook_style_save_profile`.
5. Inspect with `outlook_style_get_profile`; use `outlook_style_update_profile` for confirmed rules, preferred/forbidden phrases, or corrections.

For a normal reply, call `outlook_style_prepare_draft_context` with the current `message_id` and `store_id`. The agent should imitate tone and structure, use the current thread as factual truth, and then call the existing `outlook_create_reply_draft`. Nothing is sent automatically.

### Maintenance and rollback

```powershell
OutlookMcp.Server.exe --style-scan-status
OutlookMcp.Server.exe --style-scan
OutlookMcp.Server.exe --style-sync
OutlookMcp.Server.exe --style-rebuild-profile
OutlookMcp.Server.exe --style-scan-status --style-db-path "D:\Private\sent-email-index.db" --style-profile-path "D:\Private\writing-profile.json"
```

`--style-scan` resumes persisted folder offsets and prints progress only, never bodies. `--style-sync` processes new or modified Sent items after the initial scan. `--style-rebuild-profile` saves a deterministic statistics-only baseline for review; nuanced AI-assisted regeneration still uses the profile-dataset/save workflow. Profile replacement and updates archive the prior JSON; list and restore versions through the MCP tools.

The default local files are:

- `%AppData%\EULE\OutlookMcp\WritingStyle\sent-email-index.db`
- `%AppData%\EULE\OutlookMcp\WritingStyle\writing-profile.json`
- `%AppData%\EULE\OutlookMcp\WritingStyle\writing-profile-history.json`

Deleting the SQLite database while the server is stopped rebuilds the index from scratch on the next scan. Deleting the profile does not delete the index. If messages appear missing, first let Outlook finish IMAP synchronisation, run `--style-sync`, and inspect scan status/failure counts. The server never connects directly to IMAP or Microsoft Graph.

Privacy note: local examples returned by a style MCP tool are sent to the configured MCP client/model provider. Outputs are bounded and full cleaned context is disabled by default, but users must review their provider's privacy terms and choose `AllowedStores`/retention settings appropriate for company mail.

## Integration testing

Integration tests must use a dedicated test mailbox/profile. They never run unless explicitly enabled:

```powershell
$env:OUTLOOK_MCP_INTEGRATION = '1'
$env:OUTLOOK_MCP_TEST_QUERY = 'known unique test subject'
dotnet test .\tests\OutlookMcp.IntegrationTests\OutlookMcp.IntegrationTests.csproj
```

Draft tests additionally require `OUTLOOK_MCP_ALLOW_DRAFT_TESTS=1`, `OUTLOOK_MCP_TEST_MESSAGE_ID`, and `OUTLOOK_MCP_TEST_STORE_ID`. Attachment-save tests require `OUTLOOK_MCP_ALLOW_ATTACHMENT_TESTS=1`, `OUTLOOK_MCP_TEST_ATTACHMENT_MESSAGE_ID`, and `OUTLOOK_MCP_TEST_ATTACHMENT_STORE_ID`. The identifiers must come from this server's search output. Test drafts are intentionally left unsent for manual inspection.

Never enable integration tests against arbitrary production mail.

## Troubleshooting

### Outlook not installed or New Outlook is active

The server requires the COM registration named `Outlook.Application`, supplied by Outlook Classic. Install/repair Classic Outlook and open it manually. New Outlook alone cannot satisfy this requirement.

### Profile or MAPI is not ready

Open Outlook Classic interactively, complete account/profile prompts, dismiss modal dialogs, and confirm folders are visible. Then rerun `--diagnose`.

### IMAP messages or search results are missing

The server never bypasses Outlook. Confirm the folder has completed local synchronisation. Narrow searches with a date range, sender, subject, or explicit folder. Search applies Outlook date/unread filters first and then inspects a bounded candidate set for body/attachment-name matches; a warning reports when the scan cap was reached.

### Stale message reference

Outlook `EntryID` values can change after a move, deletion, or IMAP resynchronisation. Successful bulk moves return the new `message_id`; otherwise search for the email again and use the new `message_id` plus `store_id`.

### Outlook bitness or activation problem

Try the release matching Office bitness, repair Office COM registration, and ensure Outlook can open the intended profile under the same Windows user. No Outlook installation path is hard-coded.

### Outlook appears busy or freezes during a query

Cancel the client request, wait for Outlook synchronisation to finish, and retry with a smaller date/folder scope. A timeout stops waiting for the result; COM cannot safely abort an already-running Outlook call, so the STA worker remains serialised until Outlook returns.

### COM security prompt or modal dialog

Return to Outlook and resolve the prompt. Draft display is optional and defaults off. The server does not automate security dialogs.

### MCP client cannot launch the executable

Use an absolute path, quote/escape it according to the client format, run `--version` from PowerShell, inspect the log directory, and increase the client's startup timeout to 30 seconds if Outlook starts slowly.

## Known limitations

- Only data already synchronised by the active Outlook Classic profile can be found.
- Outlook's object model and IMAP provider determine folder names, conversation metadata, MIME metadata, and EntryID stability.
- Conversation and related-mail fallback is deterministic and bounded; it can miss messages when Outlook metadata and subjects both changed.
- Search is Outlook-filtered but body/project matching fairly samples a capped candidate set across selected folders rather than maintaining a local index. The response reports folder, scan, and truncation counts. No vector database, embeddings, Graph, or direct IMAP access is used.
- Filing-rule historical matching is representative rather than exhaustive. Outlook uses substring matching for sender-address and text conditions, and Outlook or the mail provider may classify some move rules as client-only, requiring Outlook Classic to be running.
- Writing-style sync detects new/modified items through Outlook timestamps and reconciles lightweight current EntryIDs to report missing indexed items. An IMAP EntryID change can appear as one missing old row plus one new row when Outlook exposes no reliable identity link; reconciliation never deletes local history or changes the mailbox.
- Authored-text, recurring signature, and disclaimer detection is heuristic. Quality/confidence indicators are exposed so ambiguous items have less influence and can be reviewed.
- Calendar sync copies events with their reminders, so an event can ring from both calendars when both are active in the same profile. Update detection relies on Outlook's `LastModificationTime`, so a provider resynchronisation can cause already-mirrored events to be refreshed once.
- Cancellation cannot interrupt an in-flight COM call safely; it cancels the caller's wait while the single STA worker finishes that call.
- HTML support is intentionally opt-in and should remain disabled unless required.

## Development structure

- `OutlookMcp.Server`: stdio MCP host, tool schemas, command-line diagnostics, logging.
- `OutlookMcp.Application`: validation, configuration, body processing, safe identifiers and paths.
- `OutlookMcp.Infrastructure`: private SQLite schema, checkpoints, statistics, and FTS5 retrieval.
- `OutlookMcp.OutlookInterop`: Outlook session, STA dispatcher, COM-to-DTO gateway.
- `OutlookMcp.Contracts`: bounded request/response and structured error contracts.
- `tests`: pure unit tests and explicitly gated Outlook integration tests.

Future indexing or additional write operations should remain behind application interfaces, bounded batches, dry-run or preview modes, and explicit approval controls.
