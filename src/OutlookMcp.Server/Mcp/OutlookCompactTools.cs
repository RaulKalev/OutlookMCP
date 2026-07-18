using System.ComponentModel;
using ModelContextProtocol.Server;
using OutlookMcp.Application.Abstractions;
using OutlookMcp.Application.Errors;
using OutlookMcp.Application.Services;
using OutlookMcp.Contracts;

namespace OutlookMcp.Server.Mcp;

/// <summary>
/// A deliberately small, high-frequency tool surface. Advanced filters, mailbox
/// organisation, attachment saving, and style maintenance remain in the full profile.
/// </summary>
[McpServerToolType]
public sealed class OutlookCompactTools(IOutlookGateway outlook, IExchangeCalendarGateway exchange, CalendarSyncCoordinator calendarSync, ToolExecutor executor)
{
    [McpServerTool(Name = "outlook_find_folders"), Description("Find mail folders by name or path. Read-only.")]
    public Task<ToolResponse<IReadOnlyList<FolderDto>>> FindFolders(
        [Description("Name or path fragment.")] string query,
        string? store_id = null,
        [Description("1-100; default 10.")] int max_results = 10,
        CancellationToken cancellationToken = default) =>
        executor.RunAsync(() => outlook.FindFoldersAsync(new(query, store_id, max_results, false), cancellationToken));

    [McpServerTool(Name = "outlook_search_emails"), Description("Search local Outlook mail. Returns compact references; read selected bodies in one batch. Read-only.")]
    public Task<ToolResponse<CompactSearchResultDto>> SearchEmails(
        [Description("Text matched across common email fields. Empty only with another filter.")] string query,
        string[]? store_ids = null,
        string[]? folder_ids = null,
        [Description("inbox_and_sent, inbox, sent, or all.")] string scope = "inbox_and_sent",
        string? sender = null,
        string? subject = null,
        DateTimeOffset? date_from = null,
        DateTimeOffset? date_to = null,
        bool unread_only = false,
        [Description("1-100; default 10.")] int max_results = 10,
        CancellationToken cancellationToken = default) => executor.RunAsync(async () =>
        {
            var (searchInbox, searchSent, searchAll) = ParseScope(scope);
            var result = await outlook.SearchEmailsAsync(new(
                query, store_ids, folder_ids, true, searchInbox, searchSent, searchAll,
                sender, null, subject, date_from, date_to, null, unread_only, max_results,
                "newest_first", false, "all_terms"), cancellationToken).ConfigureAwait(false);
            return CompactDtoMapper.ToCompact(result);
        });

    [McpServerTool(Name = "outlook_read_emails_batch"), Description("Read one or more emails from one store in a single compact call. Read-only.")]
    public Task<ToolResponse<CompactBatchReadResultDto>> ReadEmailsBatch(
        string[] message_ids,
        string store_id,
        [Description("Plain-text characters per email; default 4000.")] int max_body_characters = 4_000,
        CancellationToken cancellationToken = default) => executor.RunAsync(async () =>
        {
            var result = await outlook.ReadEmailsBatchAsync(
                new(message_ids, store_id, "plain_text", max_body_characters, false, true),
                cancellationToken).ConfigureAwait(false);
            return CompactDtoMapper.ToCompact(result);
        });

    [McpServerTool(Name = "outlook_get_selected_email"), Description("Read the email selected or open in Outlook. Read-only.")]
    public Task<ToolResponse<CompactSelectionDto>> GetSelectedEmail(
        bool include_body = true,
        [Description("1-100; default 5.")] int max_messages = 5,
        [Description("Characters per email; default 5000.")] int max_body_characters = 5_000,
        CancellationToken cancellationToken = default) => executor.RunAsync(async () =>
        {
            var result = await outlook.GetSelectedEmailAsync(
                new(include_body, max_messages, max_body_characters), cancellationToken).ConfigureAwait(false);
            return CompactDtoMapper.ToCompact(result);
        });

    [McpServerTool(Name = "outlook_find_related_emails"), Description("Find related emails using local deterministic signals. Read-only.")]
    public Task<ToolResponse<IReadOnlyList<CompactRelatedEmailDto>>> FindRelatedEmails(
        string message_id,
        string store_id,
        [Description("1-100; default 10.")] int max_results = 10,
        CancellationToken cancellationToken = default) => executor.RunAsync(async () =>
        {
            var result = await outlook.FindRelatedEmailsAsync(
                new(message_id, store_id, max_results), cancellationToken).ConfigureAwait(false);
            return CompactDtoMapper.ToCompact(result);
        });

    [McpServerTool(Name = "outlook_list_attachments"), Description("List attachment metadata without opening content. Read-only.")]
    public Task<ToolResponse<IReadOnlyList<AttachmentDto>>> ListAttachments(
        string message_id,
        string store_id,
        CancellationToken cancellationToken = default) =>
        executor.RunAsync(() => outlook.ListAttachmentsAsync(message_id, store_id, cancellationToken));

    [McpServerTool(Name = "outlook_list_calendars"), Description("List calendar folders per store, marking defaults. Read-only.")]
    public Task<ToolResponse<IReadOnlyList<CalendarFolderDto>>> ListCalendars(
        string? store_id = null,
        CancellationToken cancellationToken = default) =>
        executor.RunAsync(() => outlook.ListCalendarFoldersAsync(store_id, cancellationToken));

    [McpServerTool(Name = "outlook_exchange_login"), Description("Start the one-time Exchange Online device-code sign-in. Show the returned URL and code to the user.")]
    public Task<ToolResponse<ExchangeLoginDto>> ExchangeLogin(CancellationToken cancellationToken = default) =>
        executor.RunAsync(() => exchange.BeginDeviceCodeLoginAsync(cancellationToken));

    [McpServerTool(Name = "outlook_exchange_auth_status"), Description("Report the Exchange sign-in state. Read-only.")]
    public Task<ToolResponse<ExchangeAuthStatusDto>> ExchangeAuthStatus(CancellationToken cancellationToken = default) =>
        executor.RunAsync(() => exchange.GetAuthStatusAsync(cancellationToken));

    [McpServerTool(Name = "outlook_exchange_logout"), Description("Sign the Exchange account out and clear cached tokens.")]
    public Task<ToolResponse<ExchangeAuthStatusDto>> ExchangeLogout(CancellationToken cancellationToken = default) =>
        executor.RunAsync(() => exchange.LogoutAsync(cancellationToken));

    [McpServerTool(Name = "outlook_sync_calendar"), Description("One-way sync of upcoming local calendar events into the signed-in Exchange account's sync-owned calendar via Microsoft Graph. Never modifies the local calendar, never sends invitations. dry_run defaults to true; apply only after the user confirms the plan.")]
    public Task<ToolResponse<CalendarSyncResultDto>> SyncCalendar(
        string? source_calendar_folder_id = null,
        string? source_store_id = null,
        [Description("Months ahead of today; default from config (3).")] int? months_ahead = null,
        bool dry_run = true,
        CancellationToken cancellationToken = default) => executor.RunAsync(() => calendarSync.SyncAsync(
            new(source_calendar_folder_id, source_store_id, months_ahead, dry_run), cancellationToken));

    [McpServerTool(Name = "outlook_create_draft"), Description("Save a new unsent draft for user review. Never sends.")]
    public Task<ToolResponse<DraftDto>> CreateDraft(
        string subject,
        string body,
        string? to = null,
        string? cc = null,
        string? bcc = null,
        string? account_or_store_id = null,
        bool display_draft = false,
        CancellationToken cancellationToken = default) => executor.RunAsync(() => outlook.CreateDraftAsync(
            new(subject, body, "plain_text", to, cc, bcc, account_or_store_id, null, display_draft), cancellationToken));

    [McpServerTool(Name = "outlook_create_reply_draft"), Description("Save an unsent native reply draft for user review. Never sends.")]
    public Task<ToolResponse<DraftDto>> CreateReplyDraft(
        string message_id,
        string store_id,
        string body,
        bool reply_all = false,
        bool display_draft = false,
        bool include_original_message = true,
        CancellationToken cancellationToken = default) => executor.RunAsync(() => outlook.CreateReplyDraftAsync(
            new(message_id, store_id, body, reply_all, "plain_text", display_draft, include_original_message), cancellationToken));

    [McpServerTool(Name = "outlook_create_forward_draft"), Description("Save an unsent native forward draft for user review. Never sends.")]
    public Task<ToolResponse<DraftDto>> CreateForwardDraft(
        string message_id,
        string store_id,
        string? body = null,
        string? to = null,
        string? cc = null,
        bool include_attachments = true,
        bool display_draft = false,
        CancellationToken cancellationToken = default) => executor.RunAsync(() => outlook.CreateForwardDraftAsync(
            new(message_id, store_id, body, to, cc, include_attachments, display_draft), cancellationToken));

    private static (bool SearchInbox, bool SearchSent, bool SearchAll) ParseScope(string scope) =>
        scope.ToLowerInvariant() switch
        {
            "inbox_and_sent" => (true, true, false),
            "inbox" => (true, false, false),
            "sent" => (false, true, false),
            "all" => (false, false, true),
            _ => throw new OutlookMcpException(
                ErrorCodes.InvalidArgument,
                "scope must be inbox_and_sent, inbox, sent, or all.",
                "Correct scope and retry.")
        };
}
