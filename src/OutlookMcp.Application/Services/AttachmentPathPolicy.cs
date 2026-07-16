using System.Text.RegularExpressions;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Services;

public sealed class AttachmentPathPolicy
{
    private readonly string[] _allowedDirectories;

    public AttachmentPathPolicy(OutlookOptions options)
    {
        _allowedDirectories = options.AllowedAttachmentDirectories
            .Select(Environment.ExpandEnvironmentVariables)
            .Select(Path.GetFullPath)
            .Select(TrimEndingSeparator)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public string DefaultDirectory => _allowedDirectories.FirstOrDefault()
        ?? throw new OutlookMcpException(ErrorCodes.AttachmentPathNotAllowed, "No attachment directory is configured.", "Add an allowed attachment directory to config.json.");

    public string ValidateDirectory(string? requested)
    {
        var candidate = TrimEndingSeparator(Path.GetFullPath(Environment.ExpandEnvironmentVariables(requested ?? DefaultDirectory)));
        if (!_allowedDirectories.Any(allowed => candidate.Equals(allowed, StringComparison.OrdinalIgnoreCase) || candidate.StartsWith(allowed + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)))
        {
            throw new OutlookMcpException(ErrorCodes.AttachmentPathNotAllowed, "The destination is outside the configured attachment directories.", "Choose an allowed directory or update config.json.");
        }

        RejectExistingReparsePoints(candidate);

        return candidate;
    }

    public static string SanitizeFileName(string fileName)
    {
        var name = Path.GetFileName(fileName);
        foreach (var invalid in Path.GetInvalidFileNameChars()) name = name.Replace(invalid, '_');
        name = Regex.Replace(name, @"[. ]+$", string.Empty, RegexOptions.CultureInvariant);
        if (string.IsNullOrWhiteSpace(name)) name = "attachment";
        var stem = Path.GetFileNameWithoutExtension(name);
        if (ReservedWindowsNames.Contains(stem)) name = "_" + name;
        return name.Length <= 180 ? name : name[..180];
    }

    public static string GetAvailablePath(string directory, string fileName, bool overwrite)
    {
        var path = Path.Combine(directory, SanitizeFileName(fileName));
        if (overwrite || !File.Exists(path)) return path;
        var stem = Path.GetFileNameWithoutExtension(path);
        var extension = Path.GetExtension(path);
        for (var index = 1; index < 10_000; index++)
        {
            var candidate = Path.Combine(directory, $"{stem} ({index}){extension}");
            if (!File.Exists(candidate)) return candidate;
        }

        throw new IOException("Unable to allocate a unique attachment filename.");
    }

    private static string TrimEndingSeparator(string path) => Path.TrimEndingDirectorySeparator(path);

    private static readonly HashSet<string> ReservedWindowsNames = new(
        ["CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"],
        StringComparer.OrdinalIgnoreCase);

    private static void RejectExistingReparsePoints(string path)
    {
        var current = new DirectoryInfo(path);
        while (current is not null)
        {
            if (current.Exists && current.Attributes.HasFlag(FileAttributes.ReparsePoint))
            {
                throw new OutlookMcpException(ErrorCodes.AttachmentPathNotAllowed, "The destination contains a filesystem link or junction.", "Choose a normal directory inside an allowed attachment root.");
            }
            current = current.Parent;
        }
    }
}
