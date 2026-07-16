using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Services;

public static class MessageReferenceCodec
{
    private const string LegacyPrefix = "omcp1_";
    private const string CompactPrefix = "omcp2_";

    public static string Encode(OutlookItemReference reference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(reference.EntryId);
        ArgumentException.ThrowIfNullOrWhiteSpace(reference.StoreId);
        return CompactPrefix + StoreFingerprint(reference.StoreId) + "_" + ToBase64Url(Encoding.UTF8.GetBytes(reference.EntryId));
    }

    public static OutlookItemReference Decode(string? encoded)
    {
        if (encoded?.StartsWith(CompactPrefix, StringComparison.Ordinal) == true)
        {
            throw new OutlookMcpException(ErrorCodes.InvalidArgument, "The compact message identifier requires its paired store_id.", "Use message_id and store_id together exactly as returned by the server.");
        }

        return DecodeLegacy(encoded);
    }

    public static OutlookItemReference Decode(string? encoded, string expectedStoreId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(expectedStoreId);
        if (string.IsNullOrWhiteSpace(encoded) || !encoded.StartsWith(CompactPrefix, StringComparison.Ordinal))
        {
            var legacy = DecodeLegacy(encoded);
            if (!string.Equals(legacy.StoreId, expectedStoreId, StringComparison.Ordinal)) throw StoreMismatch();
            return legacy;
        }

        try
        {
            var separator = encoded.IndexOf('_', CompactPrefix.Length);
            if (separator < 0) throw new FormatException("Compact reference is missing its separator.");
            var fingerprint = encoded[CompactPrefix.Length..separator];
            if (!string.Equals(fingerprint, StoreFingerprint(expectedStoreId), StringComparison.Ordinal)) throw StoreMismatch();
            var entryId = Encoding.UTF8.GetString(FromBase64Url(encoded[(separator + 1)..]));
            if (string.IsNullOrWhiteSpace(entryId)) throw new FormatException("Compact reference has no entry identifier.");
            return new OutlookItemReference(entryId, expectedStoreId);
        }
        catch (OutlookMcpException) { throw; }
        catch (Exception ex) when (ex is FormatException or DecoderFallbackException)
        {
            throw Malformed(ex);
        }
    }

    private static OutlookItemReference DecodeLegacy(string? encoded)
    {
        if (string.IsNullOrWhiteSpace(encoded) || !encoded.StartsWith(LegacyPrefix, StringComparison.Ordinal))
        {
            throw new OutlookMcpException(ErrorCodes.InvalidArgument, "The message identifier is not a valid Outlook MCP reference.", "Search for the message again and use the returned message_id and store_id.");
        }

        try
        {
            var result = JsonSerializer.Deserialize<OutlookItemReference>(FromBase64Url(encoded[LegacyPrefix.Length..]));
            if (result is null || string.IsNullOrWhiteSpace(result.EntryId) || string.IsNullOrWhiteSpace(result.StoreId))
            {
                throw new FormatException("Reference payload is incomplete.");
            }

            return result;
        }
        catch (Exception ex) when (ex is FormatException or JsonException)
        {
            throw Malformed(ex);
        }
    }

    private static string StoreFingerprint(string storeId) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(storeId)))[..16];
    private static string ToBase64Url(byte[] value) => Convert.ToBase64String(value).TrimEnd('=').Replace('+', '-').Replace('/', '_');
    private static byte[] FromBase64Url(string value)
    {
        var payload = value.Replace('-', '+').Replace('_', '/');
        payload = payload.PadRight(payload.Length + ((4 - payload.Length % 4) % 4), '=');
        return Convert.FromBase64String(payload);
    }

    private static OutlookMcpException StoreMismatch() => new(ErrorCodes.InvalidArgument, "message_id and store_id refer to different Outlook stores.", "Use the paired message_id and store_id returned by the same result.");
    private static OutlookMcpException Malformed(Exception ex) => new(ErrorCodes.InvalidArgument, "The message identifier is malformed.", "Search for the message again and use its newly returned identifiers.", ex);
}
