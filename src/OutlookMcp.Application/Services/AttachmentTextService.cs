using System.Diagnostics;
using Microsoft.Extensions.Logging;
using OutlookMcp.Application.Abstractions;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Services;

/// <summary>
/// Reads the text of one attachment without changing Outlook: the bytes are saved into a fresh private
/// scratch folder, parsed off the Outlook STA thread under a timeout, and the folder is always deleted.
/// </summary>
public sealed class AttachmentTextService
{
    private const string ToolName = "outlook_read_attachment_text";
    private static readonly TimeSpan CancellationGrace = TimeSpan.FromSeconds(5);
    private readonly IOutlookGateway _gateway;
    private readonly IAttachmentTextExtractor _extractor;
    private readonly AttachmentScratchStore _scratch;
    private readonly AttachmentTextOptions _options;
    private readonly ILogger<AttachmentTextService> _logger;

    public AttachmentTextService(IOutlookGateway gateway, IAttachmentTextExtractor extractor, AttachmentScratchStore scratch, OutlookMcpOptions options, ILogger<AttachmentTextService> logger)
    {
        _gateway = gateway;
        _extractor = extractor;
        _scratch = scratch;
        _options = options.AttachmentText;
        _logger = logger;
    }

    public async Task<AttachmentTextDto> ReadAsync(string messageId, string storeId, int attachmentId, int maxCharacters, int pageFrom, int? pageTo, string? sheet, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(messageId) || string.IsNullOrWhiteSpace(storeId)) throw Invalid("message_id and store_id are required.");
        if (attachmentId < 1) throw Invalid("attachment_id must be at least 1.");
        if (maxCharacters < 1) throw Invalid("max_characters must be at least 1.");
        if (pageFrom < 1) throw Invalid("page_from must be at least 1.");
        if (pageTo is not null && pageTo < pageFrom) throw Invalid("page_to must be greater than or equal to page_from.");
        if (sheet is { Length: > 200 }) throw Invalid("sheet cannot exceed 200 characters.");

        var warnings = new List<string>();
        var characters = maxCharacters;
        if (characters > _options.MaximumCharacters)
        {
            characters = _options.MaximumCharacters;
            warnings.Add($"max_characters was capped at {_options.MaximumCharacters}.");
        }

        var stopwatch = Stopwatch.StartNew();
        string? extension = null;
        var folder = _scratch.CreateFolder();
        try
        {
            var file = await _gateway.SaveAttachmentToScratchAsync(
                new(messageId, storeId, attachmentId, folder, _options.MaximumFileBytes, _extractor.SupportedExtensions), cancellationToken).ConfigureAwait(false);
            extension = file.Extension;

            if (file.NotSavedReason is AttachmentScratchFileDto.UnsupportedType or AttachmentScratchFileDto.NotAFile)
            {
                var reason = file.NotSavedReason == AttachmentScratchFileDto.NotAFile
                    ? "Embedded item, link or OLE object: not a file attachment."
                    : _extractor.GetUnsupportedReason(file.Extension) ?? "The file type is not supported for text extraction.";
                _logger.LogInformation("Tool {ToolName} skipped unsupported attachment; Extension={Extension}, SizeBytes={SizeBytes}", ToolName, file.Extension, file.SizeBytes);
                return Unsupported(file, reason, warnings);
            }

            if (file.SavedPath is null || file.SizeBytes > _options.MaximumFileBytes) throw TooLarge(file.SizeBytes);

            var request = new AttachmentTextRequest(characters, pageFrom, pageTo, string.IsNullOrWhiteSpace(sheet) ? null : sheet.Trim());
            var result = await ExtractWithTimeoutAsync(file.SavedPath, file, request, cancellationToken).ConfigureAwait(false);
            if (warnings.Count > 0) result = result with { Warnings = [.. warnings, .. result.Warnings] };
            _logger.LogInformation(
                "Tool {ToolName} read attachment text; Extension={Extension}, SizeBytes={SizeBytes}, PageCount={PageCount}, SheetCount={SheetCount}, Characters={Characters}, Truncated={Truncated}, ElapsedMilliseconds={ElapsedMilliseconds}",
                ToolName, file.Extension, file.SizeBytes, result.PageCount, result.SheetNames?.Count, result.CharactersReturned, result.Truncated, stopwatch.ElapsedMilliseconds);
            return result;
        }
        catch (OutlookMcpException ex)
        {
            _logger.LogWarning("Tool {ToolName} failed with {ErrorCode}; Extension={Extension}, ElapsedMilliseconds={ElapsedMilliseconds}", ToolName, ex.Code, extension, stopwatch.ElapsedMilliseconds);
            throw;
        }
        finally
        {
            DeleteScratchFolder(folder);
        }
    }

    private async Task<AttachmentTextDto> ExtractWithTimeoutAsync(string path, AttachmentScratchFileDto file, AttachmentTextRequest request, CancellationToken cancellationToken)
    {
        using var parsing = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var task = Task.Run(() => _extractor.Extract(path, file, request, parsing.Token), CancellationToken.None);
        try
        {
            return await task.WaitAsync(TimeSpan.FromSeconds(_options.ParseTimeoutSeconds), cancellationToken).ConfigureAwait(false);
        }
        catch (TimeoutException)
        {
            await parsing.CancelAsync().ConfigureAwait(false);
            await WaitForParserToStopAsync(task).ConfigureAwait(false);
            throw new OutlookMcpException(ErrorCodes.AttachmentParseFailed, $"Reading the attachment took longer than {_options.ParseTimeoutSeconds} seconds and was stopped.",
                "Request fewer pages or one sheet, or open the file in its own application.");
        }
        catch (Exception ex) when (ex is not OutlookMcpException and not OperationCanceledException)
        {
            throw new OutlookMcpException(ErrorCodes.AttachmentParseFailed, "The attachment could not be parsed.", "The file may be damaged or use an unsupported variant; open it in its own application.", ex);
        }
    }

    private static async Task WaitForParserToStopAsync(Task task)
    {
        // The parser's own outcome no longer matters once the call has timed out; this only gives it time to release the file.
        try { await task.WaitAsync(CancellationGrace).ConfigureAwait(false); }
        catch (Exception) { }
    }

    private void DeleteScratchFolder(string folder)
    {
        try { _scratch.DeleteFolder(folder); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _logger.LogWarning("Tool {ToolName} could not delete its scratch folder yet; it is removed on the next server start", ToolName);
        }
    }

    private static AttachmentTextDto Unsupported(AttachmentScratchFileDto file, string reason, IReadOnlyList<string> warnings) => new(
        file.FileName, file.Extension, file.MimeType, file.SizeBytes, false, reason, null, null, null, null, null, 0, false, string.Empty, warnings);

    private static OutlookMcpException TooLarge(long size) => new(ErrorCodes.AttachmentTooLarge,
        $"The attachment is {size / (1024 * 1024)} MB, above the configured AttachmentText.MaximumFileBytes limit.", "Open the file in its own application, or raise the limit in config.json.");

    private static OutlookMcpException Invalid(string message) => new(ErrorCodes.InvalidArgument, message, "Correct the request parameters and retry.");
}
