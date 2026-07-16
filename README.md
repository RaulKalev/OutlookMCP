# EULE Outlook MCP

EULE Outlook MCP is a local, safety-first Model Context Protocol server for Microsoft Outlook Classic on Windows. It lets an MCP-capable AI client inspect Outlook connection status, find and create folders, search and read locally synchronised mail, inspect the current Outlook selection, find related correspondence, save explicitly selected attachments, move confirmed message batches, and create unsent drafts.

Outlook remains the mail client and synchronisation engine. The server talks only to the local Outlook Object Model through COM; it does not connect to Telia IMAP directly and has no Microsoft Graph or Exchange dependency. An IMAP mailbox works because Outlook Classic has already exposed its synchronised stores, folders, and messages through MAPI.

Only **Microsoft Outlook Classic for Windows** is supported. New Outlook, Outlook Web, macOS, and mobile Outlook are not supported.

## Safety and privacy

- There is no send-email, delete-email, archive-by-guessing, or mailbox-monitoring tool.
- Folder creation is explicit. Bulk moves require exact message identifiers and one exact destination folder identifier; `dry_run` defaults to `true`, duplicate identifiers are rejected, and real moves return fresh message identifiers.
- New, reply, reply-all, and forward operations only save drafts. Every draft result reports `sent: false`; the user must review and send it manually in Outlook.
- Email bodies are marked as untrusted external data. Instructions in email content must not be treated as trusted agent instructions.
- HTML body return and HTML drafting are disabled by default. HTML is parsed as data and is never rendered by this server.
- Search results and bodies have strict limits. Search only sees content Outlook has synchronised locally.
- Attachments are never opened, executed, uploaded, or extracted automatically. A single selected attachment can be saved only below an allow-listed directory. Path traversal, Windows device names, existing junctions/symlinks, and accidental overwrite are blocked.
- Normal logs contain tool names, durations, bounded identifiers, counts, actions, and error codes—not full bodies, recipient lists, credentials, or attachment contents.
- Email content is passed to the MCP client that invoked the tool. Review that AI provider's privacy and data-processing terms before use.

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
| `outlook_create_draft` | Creates an unsent draft |
| `outlook_create_reply_draft` | Creates an unsent Reply/Reply All draft |
| `outlook_create_forward_draft` | Creates an unsent forward draft |

Message tools require both the encoded `message_id` and its paired `store_id`. Batch tools share one `store_id` across their message list to keep requests compact. If Outlook moves or resynchronises an IMAP message, use the fresh identifier returned by `outlook_move_emails` or search again to replace a stale reference.

Search body previews are opt-in. The default compact search response is intended for candidate selection; use `outlook_read_emails_batch` only for the small set whose bodies are actually needed. Multi-word search defaults to `all_terms`, while `query_mode: "phrase"` is available for literal phrase matching. Search budgets are distributed across selected folders before results are globally sorted, preventing an early folder from monopolising the result limit.

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

### Installer

1. Run `artifacts\installer\EULE-Outlook-MCP-Setup-1.1.0.exe`.
2. No administrator elevation is required; the default location is `%LocalAppData%\Programs\EULE Outlook MCP`.
3. The installer creates Start-menu shortcuts for diagnostics and this README and registers an uninstaller.

### Portable ZIP

1. Extract `artifacts\EULE-Outlook-MCP-1.1.0-win-x64.zip` to a stable per-user folder.
2. Do not move the executable after configuring an MCP client unless you update the client command path.
3. Run the diagnostic command once before adding it to a client.

On first start the server creates:

- Configuration: `%AppData%\EULE\OutlookMcp\config.json`
- Logs: `%LocalAppData%\EULE\OutlookMcp\Logs`

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

Replace the executable path below with the installed path. Restart the MCP client after editing its configuration.

### Codex app, CLI, or IDE extension

Add the server through **Settings > MCP servers > Add server**, choose **STDIO**, and select the executable. Alternatively, add this to `~/.codex/config.toml`:

```toml
[mcp_servers.outlook]
command = 'C:\Users\YOUR_NAME\AppData\Local\Programs\EULE Outlook MCP\OutlookMcp.Server.exe'
args = []
startup_timeout_sec = 30
tool_timeout_sec = 60
enabled = true
```

Then restart Codex and use `/mcp` or `codex mcp list` to confirm the connection.

### Claude Desktop and JSON-based MCP clients

Merge this into the client's MCP configuration:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\Users\\YOUR_NAME\\AppData\\Local\\Programs\\EULE Outlook MCP\\OutlookMcp.Server.exe",
      "args": []
    }
  }
}
```

The server writes protocol messages only to stdout. Structured logs go to rolling local files, so they do not corrupt stdio transport.

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
  "Logging": {
    "Level": "Information",
    "RetentionDays": 14,
    "IncludeTechnicalDetails": false
  }
}
```

An empty `AllowedStores` list exposes all profile stores except blocked folders. To restrict stores, enter exact `store_id` values from `outlook_list_stores` or exact display names. Environment variables in attachment directories are expanded. Existing filesystem links/junctions in attachment destination paths are rejected.

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
- Cancellation cannot interrupt an in-flight COM call safely; it cancels the caller's wait while the single STA worker finishes that call.
- HTML support is intentionally opt-in and should remain disabled unless required.

## Development structure

- `OutlookMcp.Server`: stdio MCP host, tool schemas, command-line diagnostics, logging.
- `OutlookMcp.Application`: validation, configuration, body processing, safe identifiers and paths.
- `OutlookMcp.OutlookInterop`: Outlook session, STA dispatcher, COM-to-DTO gateway.
- `OutlookMcp.Contracts`: bounded request/response and structured error contracts.
- `tests`: pure unit tests and explicitly gated Outlook integration tests.

Future indexing or additional write operations should remain behind application interfaces, bounded batches, dry-run or preview modes, and explicit approval controls.
