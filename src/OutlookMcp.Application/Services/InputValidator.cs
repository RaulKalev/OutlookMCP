using System.Net.Mail;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.Application.Services;

public static class InputValidator
{
    public static SearchEmailsRequest Validate(SearchEmailsRequest request, OutlookOptions options)
    {
        if (request.MaxResults < 1 || request.MaxResults > options.MaximumSearchLimit)
        {
            throw new OutlookMcpException(ErrorCodes.ResultLimitExceeded, $"max_results must be between 1 and {options.MaximumSearchLimit}.", "Reduce max_results and retry.");
        }

        if (request.DateFrom > request.DateTo)
        {
            throw new OutlookMcpException(ErrorCodes.InvalidArgument, "date_from must not be later than date_to.", "Correct the date range and retry.");
        }

        if (request.SortOrder is not ("newest_first" or "oldest_first"))
        {
            throw new OutlookMcpException(ErrorCodes.InvalidArgument, "sort_order must be newest_first or oldest_first.", "Use a supported sort order.");
        }

        return request;
    }

    public static void ValidateBodyFormat(string bodyFormat)
    {
        if (bodyFormat is not ("plain_text" or "html" or "both"))
        {
            throw new OutlookMcpException(ErrorCodes.InvalidArgument, "body_format must be plain_text, html, or both.", "Use a supported body format.");
        }
    }

    public static void ValidateRecipients(params string?[] lists)
    {
        foreach (var list in lists.Where(value => !string.IsNullOrWhiteSpace(value)))
        {
            foreach (var part in list!.Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                try { _ = new MailAddress(part); }
                catch (FormatException ex)
                {
                    throw new OutlookMcpException(ErrorCodes.InvalidArgument, $"Invalid recipient address: {part}", "Use semicolon-separated RFC email addresses.", ex);
                }
            }
        }
    }
}
