using OutlookMcp.Application.Errors;
using OutlookMcp.Application.Services;
using OutlookMcp.Contracts;

namespace OutlookMcp.UnitTests;

public sealed class MessageReferenceCodecTests
{
    [Fact]
    public void EncodeDecode_RoundTripsUnicodeAndOpaqueIds()
    {
        var reference = new OutlookItemReference("entry/+ Õ", "store==/Ä");
        var encoded = MessageReferenceCodec.Encode(reference);
        var decoded = MessageReferenceCodec.Decode(encoded);
        Assert.StartsWith("omcp1_", encoded, StringComparison.Ordinal);
        Assert.Equal(reference, decoded);
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
