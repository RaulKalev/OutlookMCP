using System.Text;
using System.Text.Json;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Services;

public static class MessageReferenceCodec
{
    private const string Prefix = "omcp1_";

    public static string Encode(OutlookItemReference reference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(reference.EntryId);
        ArgumentException.ThrowIfNullOrWhiteSpace(reference.StoreId);
        var bytes = JsonSerializer.SerializeToUtf8Bytes(reference);
        return Prefix + Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');
    }

    public static OutlookItemReference Decode(string encoded)
    {
        if (string.IsNullOrWhiteSpace(encoded) || !encoded.StartsWith(Prefix, StringComparison.Ordinal))
        {
            throw new OutlookMcpException(ErrorCodes.InvalidArgument, "The message identifier is not a valid Outlook MCP reference.", "Search for the message again and use the returned message_id and store_id.");
        }

        try
        {
            var payload = encoded[Prefix.Length..].Replace('-', '+').Replace('_', '/');
            payload = payload.PadRight(payload.Length + ((4 - payload.Length % 4) % 4), '=');
            var result = JsonSerializer.Deserialize<OutlookItemReference>(Convert.FromBase64String(payload));
            if (result is null || string.IsNullOrWhiteSpace(result.EntryId) || string.IsNullOrWhiteSpace(result.StoreId))
            {
                throw new FormatException("Reference payload is incomplete.");
            }

            return result;
        }
        catch (Exception ex) when (ex is FormatException or JsonException)
        {
            throw new OutlookMcpException(ErrorCodes.InvalidArgument, "The message identifier is malformed.", "Search for the message again and use its newly returned identifiers.", ex);
        }
    }
}
