using Microsoft.Extensions.Logging.Abstractions;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Services;
using OutlookMcp.Contracts;
using OutlookMcp.OutlookInterop;

namespace OutlookMcp.IntegrationTests;

public sealed class OutlookIntegrationTests
{
    [OutlookFact]
    public async Task ConnectListStoresFoldersAndSelection()
    {
        await using var gateway = CreateGateway();
        var status = await gateway.GetStatusAsync(CancellationToken.None);
        Assert.True(status.OutlookClassicInstalled);
        Assert.True(status.MapiAvailable);
        var stores = await gateway.ListStoresAsync(CancellationToken.None);
        Assert.NotEmpty(stores);
        var folders = await gateway.ListFoldersAsync(new ListFoldersRequest(stores[0].StoreId), CancellationToken.None);
        Assert.NotEmpty(folders);
        _ = await gateway.GetSelectedEmailAsync(new SelectedEmailRequest(false, 1, 1), CancellationToken.None);
    }

    [OutlookFact]
    public async Task SearchAndReadKnownTestEmail()
    {
        var query = Required("OUTLOOK_MCP_TEST_QUERY");
        await using var gateway = CreateGateway();
        var results = await gateway.SearchEmailsAsync(new SearchEmailsRequest(query, MaxResults: 5), CancellationToken.None);
        Assert.NotEmpty(results.Messages);
        var message = results.Messages[0];
        var detail = await gateway.ReadEmailAsync(new ReadEmailRequest(message.MessageId, message.StoreId, MaxBodyCharacters: 2_000), CancellationToken.None);
        Assert.Equal(message.MessageId, detail.MessageId);
    }

    [OutlookFact]
    public async Task DiscoverSentFoldersAndReadBoundedBatchWithoutMailboxChanges()
    {
        await using var gateway = CreateGateway();
        var folders = await gateway.DiscoverSentFoldersAsync(CancellationToken.None);
        Assert.NotEmpty(folders);
        var folder = folders.First(value => value.TotalItems > 0);
        var batch = await gateway.ReadSentFolderBatchAsync(folder.StoreId, folder.FolderId, 0, 2, null, CancellationToken.None);
        Assert.InRange(batch.Messages.Count, 1, 2);
        Assert.Equal(0, batch.StartOffset);
        Assert.All(batch.Messages, message => Assert.Equal(folder.FolderId, message.FolderId));
    }

    [OutlookFact(writesDraft: true)]
    public async Task CreateNewReplyAndForwardDraftsRemainUnsent()
    {
        await using var gateway = CreateGateway();
        var draft = await gateway.CreateDraftAsync(new CreateDraftRequest($"Outlook MCP integration {Guid.NewGuid():N}", "Test draft — ära saada."), CancellationToken.None);
        Assert.False(draft.Sent);
        var reference = Required("OUTLOOK_MCP_TEST_MESSAGE_ID");
        var store = Required("OUTLOOK_MCP_TEST_STORE_ID");
        var reply = await gateway.CreateReplyDraftAsync(new CreateReplyDraftRequest(reference, store, "Integration test reply — do not send."), CancellationToken.None);
        var forward = await gateway.CreateForwardDraftAsync(new CreateForwardDraftRequest(reference, store, "Integration test forward — do not send."), CancellationToken.None);
        Assert.False(reply.Sent);
        Assert.False(forward.Sent);
    }

    [OutlookFact(writesFile: true)]
    public async Task ListAndSaveKnownSafeAttachment()
    {
        var reference = Required("OUTLOOK_MCP_TEST_ATTACHMENT_MESSAGE_ID");
        var store = Required("OUTLOOK_MCP_TEST_ATTACHMENT_STORE_ID");
        await using var gateway = CreateGateway();
        var attachments = await gateway.ListAttachmentsAsync(reference, store, CancellationToken.None);
        Assert.NotEmpty(attachments);
        var saved = await gateway.SaveAttachmentAsync(new SaveAttachmentRequest(reference, store, attachments[0].AttachmentId), CancellationToken.None);
        Assert.True(File.Exists(saved.Path));
    }

    private static OutlookGateway CreateGateway()
    {
        var downloadRoot = Path.Combine(Path.GetTempPath(), "OutlookMcpIntegration");
        var options = new OutlookMcpOptions
        {
            Outlook = new OutlookOptions { StartIfNotRunning = false, AllowedAttachmentDirectories = [downloadRoot] }
        };
        var dispatcher = new OutlookStaDispatcher(options, NullLogger<OutlookStaDispatcher>.Instance);
        return new OutlookGateway(dispatcher, options, new EmailBodyCleaner(), NullLogger<OutlookGateway>.Instance);
    }

    private static string Required(string name) => Environment.GetEnvironmentVariable(name) ?? throw new InvalidOperationException($"Set {name} for this dedicated-profile integration test.");
}
