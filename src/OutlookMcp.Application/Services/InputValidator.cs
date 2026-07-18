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

        if (request.QueryMode is not ("all_terms" or "phrase"))
        {
            throw new OutlookMcpException(ErrorCodes.InvalidArgument, "query_mode must be all_terms or phrase.", "Use a supported query mode.");
        }

        return request;
    }

    public static void ValidateBatch(IReadOnlyList<string> messageIds, int maximumBatchSize)
    {
        if (messageIds.Count < 1 || messageIds.Count > maximumBatchSize)
        {
            throw new OutlookMcpException(ErrorCodes.ResultLimitExceeded, $"message_ids must contain between 1 and {maximumBatchSize} items.", "Reduce the batch size and retry.");
        }

        if (messageIds.Any(string.IsNullOrWhiteSpace)) throw Invalid("message_ids cannot contain empty values.");
        if (messageIds.Distinct(StringComparer.Ordinal).Count() != messageIds.Count) throw Invalid("message_ids cannot contain duplicates.");
    }

    public static void ValidateFolderName(string displayName)
    {
        if (string.IsNullOrWhiteSpace(displayName)) throw Invalid("display_name is required.");
        if (displayName.Length > 255) throw Invalid("display_name must not exceed 255 characters.");
        if (displayName.IndexOfAny(['\\', '\r', '\n', '\0']) >= 0) throw Invalid("display_name contains an unsupported character.");
        if (!string.Equals(displayName, displayName.Trim(), StringComparison.Ordinal)) throw Invalid("display_name cannot start or end with whitespace.");
    }

    public static CreateFolderRuleRequest ValidateFolderRule(CreateFolderRuleRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.StoreId)) throw Invalid("store_id is required.");
        if (string.IsNullOrWhiteSpace(request.DestinationFolderId)) throw Invalid("destination_folder_id is required.");
        if (string.IsNullOrWhiteSpace(request.RuleName)) throw Invalid("rule_name is required.");
        if (request.RuleName.Length > 255) throw Invalid("rule_name must not exceed 255 characters.");
        if (!string.Equals(request.RuleName, request.RuleName.Trim(), StringComparison.Ordinal)) throw Invalid("rule_name cannot start or end with whitespace.");
        if (request.RuleName.IndexOfAny(['\r', '\n', '\0']) >= 0) throw Invalid("rule_name contains an unsupported character.");

        var groups = new[] { request.SenderAddressContains, request.SubjectContains, request.BodyContains, request.BodyOrSubjectContains };
        if (groups.All(value => value is null || value.Count == 0)) throw Invalid("At least one rule condition is required.");
        foreach (var group in groups) ValidateRuleValues(group);
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

    private static OutlookMcpException Invalid(string message) => new(ErrorCodes.InvalidArgument, message, "Correct the request parameters and retry.");

    private static void ValidateRuleValues(IReadOnlyList<string>? values)
    {
        if (values is null) return;
        if (values.Count > 20) throw Invalid("Each rule condition can contain at most 20 values.");
        if (values.Any(string.IsNullOrWhiteSpace)) throw Invalid("Rule condition values cannot be empty.");
        if (values.Any(value => value.Length > 255 || value.IndexOfAny(['\r', '\n', '\0']) >= 0)) throw Invalid("Rule condition values must not exceed 255 characters or contain control characters.");
        if (values.Any(value => !string.Equals(value, value.Trim(), StringComparison.Ordinal))) throw Invalid("Rule condition values cannot start or end with whitespace.");
        if (values.Distinct(StringComparer.OrdinalIgnoreCase).Count() != values.Count) throw Invalid("Rule condition values cannot contain duplicates.");
    }
}
