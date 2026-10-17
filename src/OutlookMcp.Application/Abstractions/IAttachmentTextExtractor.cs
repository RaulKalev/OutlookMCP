using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Abstractions;

public interface IAttachmentTextExtractor
{
    IReadOnlySet<string> SupportedExtensions { get; }

    /// <summary>Returns a reason when the extension is never read, otherwise null.</summary>
    string? GetUnsupportedReason(string extension);

    AttachmentTextDto Extract(string path, AttachmentScratchFileDto file, AttachmentTextRequest request, CancellationToken cancellationToken);
}
