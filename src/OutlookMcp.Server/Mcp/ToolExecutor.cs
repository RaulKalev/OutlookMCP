using Microsoft.Extensions.Logging;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;

namespace OutlookMcp.Server.Mcp;

public sealed class ToolExecutor(OutlookMcpOptions options, ILogger<ToolExecutor> logger)
{
    public async Task<ToolResponse<T>> RunAsync<T>(Func<Task<T>> operation)
    {
        try { return ToolResponse.Ok(await operation().ConfigureAwait(false)); }
        catch (OutlookMcpException ex) { return ToolResponse.Fail<T>(ex.ToError(options.Logging.IncludeTechnicalDetails)); }
        catch (Exception ex)
        {
            logger.LogError(ex, "Unexpected MCP tool failure");
            return ToolResponse.Fail<T>(new ErrorDto(ErrorCodes.ComOperationFailed, "The local Outlook MCP server encountered an unexpected error.", "Retry once, then run the server with --diagnose.", options.Logging.IncludeTechnicalDetails ? ex.Message : null));
        }
    }
}
