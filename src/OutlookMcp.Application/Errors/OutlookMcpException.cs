using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Errors;

public sealed class OutlookMcpException : Exception
{
    public OutlookMcpException(string code, string message, string recovery, Exception? innerException = null)
        : base(message, innerException)
    {
        Code = code;
        Recovery = recovery;
    }

    public string Code { get; }
    public string Recovery { get; }

    public ErrorDto ToError(bool includeTechnicalDetails) => new(
        Code,
        Message,
        Recovery,
        includeTechnicalDetails ? InnerException?.Message : null);
}
