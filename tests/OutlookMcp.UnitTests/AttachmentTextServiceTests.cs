using System.Text;
using Microsoft.Extensions.Logging.Abstractions;
using OutlookMcp.Application.Abstractions;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Application.Services;
using OutlookMcp.Contracts;
using OutlookMcp.Infrastructure.Attachments;

namespace OutlookMcp.UnitTests;

public sealed class AttachmentTextServiceTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "outlook-mcp-scratch-tests", Guid.NewGuid().ToString("N"));

    [Fact]
    public async Task Success_DeletesScratchFolder()
    {
        var gateway = new ScratchGateway("protokoll.txt", Encoding.UTF8.GetBytes("Koosoleku protokoll"));
        var service = CreateService(gateway);

        var result = await service.ReadAsync("message", "store", 1, 20_000, 1, null, null, CancellationToken.None);

        Assert.Equal("Koosoleku protokoll", result.Text);
        Assert.Equal("protokoll.txt", result.FileName);
        Assert.NotNull(gateway.LastDirectory);
        Assert.False(Directory.Exists(gateway.LastDirectory));
        Assert.Empty(Directory.EnumerateFileSystemEntries(_root));
    }

    [Fact]
    public async Task ParseError_DeletesScratchFolder()
    {
        var gateway = new ScratchGateway("katki.pdf", Encoding.ASCII.GetBytes("this is not a pdf"));
        var service = CreateService(gateway);

        var error = await Assert.ThrowsAsync<OutlookMcpException>(() => service.ReadAsync("message", "store", 1, 20_000, 1, null, null, CancellationToken.None));

        Assert.Equal(ErrorCodes.AttachmentParseFailed, error.Code);
        Assert.Empty(Directory.EnumerateFileSystemEntries(_root));
    }

    [Fact]
    public async Task Timeout_StopsParserAndDeletesScratchFolder()
    {
        var gateway = new ScratchGateway("aeglane.txt", [1, 2, 3]);
        var service = CreateService(gateway, new BlockingExtractor(), options => options.ParseTimeoutSeconds = 1);

        var error = await Assert.ThrowsAsync<OutlookMcpException>(() => service.ReadAsync("message", "store", 1, 20_000, 1, null, null, CancellationToken.None));

        Assert.Equal(ErrorCodes.AttachmentParseFailed, error.Code);
        Assert.Contains("longer than 1 seconds", error.Message, StringComparison.Ordinal);
        Assert.Empty(Directory.EnumerateFileSystemEntries(_root));
    }

    [Fact]
    public async Task UnsupportedType_IsNotSavedAndReturnsMetadata()
    {
        var gateway = new ScratchGateway("mudel.ifc", [1, 2, 3]);
        var service = CreateService(gateway);

        var result = await service.ReadAsync("message", "store", 1, 20_000, 1, null, null, CancellationToken.None);

        Assert.False(result.Supported);
        Assert.Contains("BIM", result.UnsupportedReason, StringComparison.Ordinal);
        Assert.Equal(3, result.SizeBytes);
        Assert.False(gateway.Saved);
        Assert.Empty(Directory.EnumerateFileSystemEntries(_root));
    }

    [Fact]
    public async Task OversizedAttachment_IsRefusedWithoutSaving()
    {
        var gateway = new ScratchGateway("suur.pdf", new byte[4_096]);
        var service = CreateService(gateway, configure: options => options.MaximumFileBytes = 1_024);

        var error = await Assert.ThrowsAsync<OutlookMcpException>(() => service.ReadAsync("message", "store", 1, 20_000, 1, null, null, CancellationToken.None));

        Assert.Equal(ErrorCodes.AttachmentTooLarge, error.Code);
        Assert.False(gateway.Saved);
        Assert.Empty(Directory.EnumerateFileSystemEntries(_root));
    }

    [Fact]
    public async Task MaximumCharacters_IsCappedByConfiguration()
    {
        var gateway = new ScratchGateway("pikk.txt", Encoding.UTF8.GetBytes(new string('a', 5_000)));
        var service = CreateService(gateway, configure: options => options.MaximumCharacters = 1_000);

        var result = await service.ReadAsync("message", "store", 1, 20_000, 1, null, null, CancellationToken.None);

        Assert.Equal(1_000, result.CharactersReturned);
        Assert.True(result.Truncated);
        Assert.Contains(result.Warnings, warning => warning.Contains("capped at 1000", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(0, 20_000, 1, null)]
    [InlineData(1, 0, 1, null)]
    [InlineData(1, 20_000, 0, null)]
    [InlineData(1, 20_000, 3, 2)]
    public async Task InvalidArguments_AreRejectedBeforeOutlookIsTouched(int attachmentId, int maxCharacters, int pageFrom, int? pageTo)
    {
        var gateway = new ScratchGateway("x.txt", [1]);
        var service = CreateService(gateway);

        var error = await Assert.ThrowsAsync<OutlookMcpException>(() => service.ReadAsync("message", "store", attachmentId, maxCharacters, pageFrom, pageTo, null, CancellationToken.None));

        Assert.Equal(ErrorCodes.InvalidArgument, error.Code);
        Assert.Null(gateway.LastDirectory);
    }

    [Fact]
    public void Scratch_DeleteStaleRemovesOnlyOldGuidFolders()
    {
        var store = new AttachmentScratchStore(_root);
        var old = store.CreateFolder();
        var fresh = store.CreateFolder();
        var foreign = Directory.CreateDirectory(Path.Combine(_root, "not-a-scratch-folder")).FullName;
        Directory.SetCreationTimeUtc(old, DateTime.UtcNow.AddHours(-2));
        Directory.SetCreationTimeUtc(foreign, DateTime.UtcNow.AddHours(-2));

        Assert.Equal(1, store.DeleteStale(TimeSpan.FromHours(1)));
        Assert.False(Directory.Exists(old));
        Assert.True(Directory.Exists(fresh));
        Assert.True(Directory.Exists(foreign));
    }

    [Fact]
    public void Scratch_RejectsFoldersOutsideTheRoot()
    {
        var store = new AttachmentScratchStore(_root);
        var outside = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));

        Assert.False(store.IsScratchFolder(outside));
        Assert.False(store.IsScratchFolder(_root));
        Assert.Throws<OutlookMcpException>(() => store.ValidateFolder(outside));
        Assert.Throws<InvalidOperationException>(() => store.DeleteFolder(outside));
    }

    private AttachmentTextService CreateService(IOutlookGateway gateway, IAttachmentTextExtractor? extractor = null, Action<AttachmentTextOptions>? configure = null)
    {
        Directory.CreateDirectory(_root);
        var options = new OutlookMcpOptions();
        configure?.Invoke(options.AttachmentText);
        return new AttachmentTextService(gateway, extractor ?? new AttachmentTextExtractor(options), new AttachmentScratchStore(_root), options, NullLogger<AttachmentTextService>.Instance);
    }

    public void Dispose()
    {
        if (Directory.Exists(_root)) Directory.Delete(_root, true);
    }

    private sealed class ScratchGateway(string fileName, byte[] content) : StubOutlookGateway
    {
        public string? LastDirectory { get; private set; }
        public bool Saved { get; private set; }

        public override Task<AttachmentScratchFileDto> SaveAttachmentToScratchAsync(AttachmentScratchRequest request, CancellationToken cancellationToken)
        {
            LastDirectory = request.ScratchDirectory;
            var extension = Path.GetExtension(fileName).ToLowerInvariant();
            if (!request.SupportedExtensions.Contains(extension))
                return Task.FromResult(new AttachmentScratchFileDto(fileName, extension, null, content.Length, null, AttachmentScratchFileDto.UnsupportedType));
            if (content.Length > request.MaximumFileBytes)
                return Task.FromResult(new AttachmentScratchFileDto(fileName, extension, null, content.Length, null, AttachmentScratchFileDto.TooLarge));
            var path = Path.Combine(request.ScratchDirectory, "attachment" + extension);
            File.WriteAllBytes(path, content);
            Saved = true;
            return Task.FromResult(new AttachmentScratchFileDto(fileName, extension, null, content.Length, path, null));
        }
    }

    private sealed class BlockingExtractor : IAttachmentTextExtractor
    {
        public IReadOnlySet<string> SupportedExtensions { get; } = new HashSet<string>([".txt"], StringComparer.OrdinalIgnoreCase);
        public string? GetUnsupportedReason(string extension) => null;

        public AttachmentTextDto Extract(string path, AttachmentScratchFileDto file, AttachmentTextRequest request, CancellationToken cancellationToken)
        {
            using var handle = File.OpenRead(path);
            while (!cancellationToken.IsCancellationRequested) Thread.Sleep(10);
            cancellationToken.ThrowIfCancellationRequested();
            throw new InvalidOperationException("unreachable");
        }
    }
}
