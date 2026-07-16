using System.Text.Json;
using OutlookMcp.Application.Errors;
using OutlookMcp.Application.Services;
using OutlookMcp.Contracts;

namespace OutlookMcp.UnitTests;

public sealed class MessageReferenceCodecTests
{
    [Fact]
    public void EncodeDecode_RoundTripsUnicodeAndOpaqueIdsWithPairedStore()
    {
        var reference = new OutlookItemReference("entry/+ Õ", "store==/Ä");
        var encoded = MessageReferenceCodec.Encode(reference);
        var decoded = MessageReferenceCodec.Decode(encoded, reference.StoreId);
        Assert.StartsWith("omcp2_", encoded, StringComparison.Ordinal);
        Assert.Equal(reference, decoded);
    }

    [Fact]
    public void Decode_RejectsMismatchedCompactStore()
    {
        var encoded = MessageReferenceCodec.Encode(new OutlookItemReference("entry", "store-one"));
        var exception = Assert.Throws<OutlookMcpException>(() => MessageReferenceCodec.Decode(encoded, "store-two"));
        Assert.Equal(ErrorCodes.InvalidArgument, exception.Code);
    }

    [Fact]
    public void Decode_AcceptsLegacyReferences()
    {
        var reference = new OutlookItemReference("legacy-entry", "legacy-store");
        var payload = Convert.ToBase64String(JsonSerializer.SerializeToUtf8Bytes(reference)).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        var encoded = "omcp1_" + payload;
        Assert.Equal(reference, MessageReferenceCodec.Decode(encoded));
        Assert.Equal(reference, MessageReferenceCodec.Decode(encoded, reference.StoreId));
    }

    [Theory]
    [InlineData("")]
    [InlineData("raw-outlook-entry-id")]
    [InlineData("omcp1_not-base64!")]
    public void Decode_RejectsMalformedReferences(string value)
    {
        var exception = Assert.Throws<OutlookMcpException>(() => MessageReferenceCodec.Decode(value));
        Assert.Equal(ErrorCodes.InvalidArgument, exception.Code);
    }
}
