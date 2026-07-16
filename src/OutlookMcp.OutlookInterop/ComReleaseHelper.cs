using System.Runtime.InteropServices;

namespace OutlookMcp.OutlookInterop;

internal static class ComReleaseHelper
{
    public static void FinalRelease(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { _ = Marshal.FinalReleaseComObject(value); }
        catch (InvalidComObjectException) { }
    }
}
