using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Services;

/// <summary>
/// Owns the private per-call folders that attachment bytes pass through while their text is read.
/// Only direct children of the root named with a GUID are ever created or deleted.
/// </summary>
public sealed class AttachmentScratchStore
{
    public AttachmentScratchStore(string root)
    {
        Root = Path.TrimEndingDirectorySeparator(Path.GetFullPath(Environment.ExpandEnvironmentVariables(root)));
    }

    public string Root { get; }

    public string CreateFolder()
    {
        Directory.CreateDirectory(Root);
        AttachmentPathPolicy.RejectExistingReparsePoints(Root);
        var folder = Path.Combine(Root, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        return folder;
    }

    public bool IsScratchFolder(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return false;
        var full = Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
        var parent = Path.GetDirectoryName(full);
        return parent is not null
            && parent.Equals(Root, StringComparison.OrdinalIgnoreCase)
            && Guid.TryParseExact(Path.GetFileName(full), "N", out _);
    }

    public string ValidateFolder(string path)
    {
        if (!IsScratchFolder(path) || !Directory.Exists(path))
        {
            throw new OutlookMcpException(ErrorCodes.AttachmentPathNotAllowed, "The attachment scratch folder is not valid.", "Retry; a fresh private scratch folder is created for every call.");
        }

        var full = Path.GetFullPath(path);
        AttachmentPathPolicy.RejectExistingReparsePoints(full);
        return full;
    }

    public void DeleteFolder(string path)
    {
        if (!IsScratchFolder(path)) throw new InvalidOperationException("Refusing to delete a folder outside the attachment scratch root.");
        if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
    }

    /// <summary>Deletes scratch folders left behind by an earlier crash; returns how many were removed.</summary>
    public int DeleteStale(TimeSpan olderThan)
    {
        if (!Directory.Exists(Root)) return 0;
        var cutoff = DateTime.UtcNow - olderThan;
        var removed = 0;
        foreach (var folder in Directory.EnumerateDirectories(Root))
        {
            if (!IsScratchFolder(folder) || Directory.GetCreationTimeUtc(folder) > cutoff) continue;
            try
            {
                Directory.Delete(folder, recursive: true);
                removed++;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Still locked by another process; the next start retries.
            }
        }

        return removed;
    }
}
