using System.Text.Json.Serialization;

namespace OutlookMcp.Contracts;

public static class ErrorCodes
{
    public const string OutlookNotInstalled = "OUTLOOK_NOT_INSTALLED";
    public const string OutlookNotAvailable = "OUTLOOK_NOT_AVAILABLE";
    public const string MapiNotReady = "MAPI_NOT_READY";
    public const string StoreNotFound = "STORE_NOT_FOUND";
    public const string FolderNotFound = "FOLDER_NOT_FOUND";
    public const string FolderAlreadyExists = "FOLDER_ALREADY_EXISTS";
    public const string MessageNotFound = "MESSAGE_NOT_FOUND";
    public const string StaleMessageReference = "STALE_MESSAGE_REFERENCE";
    public const string UnsupportedOutlookItem = "UNSUPPORTED_OUTLOOK_ITEM";
    public const string SearchTimeout = "SEARCH_TIMEOUT";
    public const string ComOperationFailed = "COM_OPERATION_FAILED";
    public const string AttachmentNotFound = "ATTACHMENT_NOT_FOUND";
    public const string AttachmentPathNotAllowed = "ATTACHMENT_PATH_NOT_ALLOWED";
    public const string ResultLimitExceeded = "RESULT_LIMIT_EXCEEDED";
    public const string InvalidArgument = "INVALID_ARGUMENT";
    public const string OperationCancelled = "OPERATION_CANCELLED";
}

public sealed record ErrorDto(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("recovery")] string Recovery,
    [property: JsonPropertyName("technical_details")] string? TechnicalDetails = null);

public sealed record ToolResponse<T>(
    [property: JsonPropertyName("success")] bool Success,
    [property: JsonPropertyName("data")] T? Data,
    [property: JsonPropertyName("error")] ErrorDto? Error);

public static class ToolResponse
{
    public static ToolResponse<T> Ok<T>(T data) => new(true, data, null);
    public static ToolResponse<T> Fail<T>(ErrorDto error) => new(false, default, error);
}
