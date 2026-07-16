namespace OutlookMcp.Contracts;

public sealed record ListFoldersRequest(string? StoreId = null, string? ParentFolderId = null, bool Recursive = false, int MaxDepth = 3, bool IncludeHidden = false);

public sealed record SearchEmailsRequest(
    string Query,
    IReadOnlyList<string>? StoreIds = null,
    IReadOnlyList<string>? FolderIds = null,
    bool IncludeSubfolders = true,
    bool SearchInbox = true,
    bool SearchSent = true,
    bool SearchAllMailFolders = false,
    string? Sender = null,
    string? Recipients = null,
    string? Subject = null,
    DateTimeOffset? DateFrom = null,
    DateTimeOffset? DateTo = null,
    bool? HasAttachments = null,
    bool UnreadOnly = false,
    int MaxResults = 25,
    string SortOrder = "newest_first",
    bool IncludeBodyPreview = true);

public sealed record ReadEmailRequest(string MessageId, string StoreId, string BodyFormat = "plain_text", int MaxBodyCharacters = 50_000, bool IncludeAttachmentMetadata = true);
public sealed record ReadThreadRequest(string MessageId, string StoreId, int MaxMessages = 30, bool IncludeSentItems = true, bool IncludeReceivedItems = true, int MaxCharactersPerMessage = 25_000);
public sealed record SelectedEmailRequest(bool IncludeBody = true, int MaxMessages = 10, int MaxBodyCharacters = 50_000);
public sealed record RelatedEmailsRequest(string MessageId, string StoreId, int MaxResults = 25, int DateRangeDays = 365, bool IncludeSameConversation = true, bool IncludeSubjectMatches = true, bool IncludeParticipantMatches = true, bool IncludeProjectKeywordMatches = true);
public sealed record SaveAttachmentRequest(string MessageId, string StoreId, int AttachmentId, string? DestinationDirectory = null, bool Overwrite = false);
public sealed record CreateDraftRequest(string Subject, string Body, string BodyFormat = "plain_text", string? To = null, string? Cc = null, string? Bcc = null, string? AccountOrStoreId = null, string? Importance = null, bool DisplayDraft = false);
public sealed record CreateReplyDraftRequest(string MessageId, string StoreId, string Body, bool ReplyAll = false, string BodyFormat = "plain_text", bool DisplayDraft = false, bool IncludeOriginalMessage = true);
public sealed record CreateForwardDraftRequest(string MessageId, string StoreId, string? Body = null, string? To = null, string? Cc = null, bool IncludeAttachments = true, bool DisplayDraft = false);
