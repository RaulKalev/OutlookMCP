using System.Globalization;
using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OutlookMcp.Application.Abstractions;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using UglyToad.PdfPig.Exceptions;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace OutlookMcp.Infrastructure.Attachments;

/// <summary>
/// Pure in-process text extraction for attachment files. Never launches another program and never writes files.
/// Every parser is bounded by characters, pages, sheets or cells and checks the cancellation token as it goes.
/// </summary>
public sealed class AttachmentTextExtractor : IAttachmentTextExtractor
{
    private const string ScannedPdfWarning = "scanned PDF: no text layer; OCR not supported";
    private static readonly HashSet<string> Supported = new([".pdf", ".xlsx", ".xlsm", ".docx", ".csv", ".txt"], StringComparer.OrdinalIgnoreCase);
    private static readonly Encoding StrictUtf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
    private readonly AttachmentTextOptions _options;

    static AttachmentTextExtractor() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    public AttachmentTextExtractor(OutlookMcpOptions options) => _options = options.AttachmentText;

    public IReadOnlySet<string> SupportedExtensions => Supported;

    public string? GetUnsupportedReason(string extension)
    {
        var value = (extension ?? string.Empty).ToLowerInvariant();
        if (Supported.Contains(value)) return null;
        return value switch
        {
            ".dwg" or ".dxf" or ".dgn" => "CAD drawing: only metadata is returned.",
            ".ifc" or ".rvt" or ".rfa" or ".nwd" or ".nwc" => "BIM model: only metadata is returned.",
            ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" or ".tif" or ".tiff" or ".webp" or ".heic" or ".emf" or ".wmf" => "Image: OCR is not supported.",
            ".zip" or ".7z" or ".rar" or ".gz" or ".tar" => "Archive: archives are never expanded.",
            ".msg" or ".eml" => "Attached email item: not read as an attachment.",
            ".doc" or ".xls" or ".ppt" => "Legacy binary Office format: only .docx and .xlsx/.xlsm are read.",
            ".pptx" => "Presentation: not supported for text extraction.",
            ".exe" or ".dll" or ".msi" or ".bat" or ".cmd" or ".ps1" or ".vbs" or ".js" or ".lnk" or ".scr" => "Executable or script: never opened.",
            "" => "The attachment has no file extension, so its type is unknown.",
            _ => "The file type is not supported for text extraction."
        };
    }

    public AttachmentTextDto Extract(string path, AttachmentScratchFileDto file, AttachmentTextRequest request, CancellationToken cancellationToken)
    {
        var size = new FileInfo(path).Length;
        if (size > _options.MaximumFileBytes)
        {
            throw new OutlookMcpException(ErrorCodes.AttachmentTooLarge, "The attachment is above the configured AttachmentText.MaximumFileBytes limit.", "Open the file in its own application, or raise the limit in config.json.");
        }

        var described = file with { SizeBytes = size };
        try
        {
            return described.Extension.ToLowerInvariant() switch
            {
                ".pdf" => ReadPdf(path, described, request, cancellationToken),
                ".xlsx" or ".xlsm" => ReadWorkbook(path, described, request, cancellationToken),
                ".docx" => ReadDocument(path, described, request, cancellationToken),
                ".csv" or ".txt" => ReadPlainText(path, described, request),
                _ => new AttachmentTextDto(described.FileName, described.Extension, described.MimeType, size, false,
                    GetUnsupportedReason(described.Extension), null, null, null, null, null, 0, false, string.Empty, [])
            };
        }
        catch (Exception ex) when (ex is not OutlookMcpException and not OperationCanceledException)
        {
            throw new OutlookMcpException(ErrorCodes.AttachmentParseFailed, "The attachment could not be parsed.", "The file may be damaged or use an unsupported variant; open it in its own application.", ex);
        }
    }

    private AttachmentTextDto ReadPdf(string path, AttachmentScratchFileDto file, AttachmentTextRequest request, CancellationToken cancellationToken)
    {
        var warnings = new List<string>();
        PdfDocument document;
        try
        {
            document = PdfDocument.Open(path, new ParsingOptions { UseLenientParsing = true, SkipMissingFonts = true });
        }
        catch (PdfDocumentEncryptedException)
        {
            warnings.Add("encrypted PDF: a password is required, so no text was extracted");
            return Result(file, null, null, null, null, null, new BoundedText(request.MaximumCharacters), false, warnings);
        }

        using (document)
        {
            var pageCount = document.NumberOfPages;
            if (document.IsEncrypted) warnings.Add("encrypted PDF: opened without a password; some text may be missing");
            var text = new BoundedText(request.MaximumCharacters);
            if (request.PageFrom > pageCount)
            {
                warnings.Add($"page_from {request.PageFrom} is after the last page ({pageCount}).");
                return Result(file, pageCount, null, null, null, null, text, false, warnings);
            }

            var lastRequested = Math.Min(request.PageTo ?? pageCount, pageCount);
            var lastParsed = Math.Min(lastRequested, request.PageFrom + _options.MaximumPdfPages - 1);
            var truncated = false;
            if (lastParsed < lastRequested)
            {
                truncated = true;
                warnings.Add($"At most {_options.MaximumPdfPages} pages are parsed per call; continue with page_from={lastParsed + 1}.");
            }

            int? returnedTo = null;
            var pagesWithText = 0;
            var pagesWithoutText = 0;
            for (var number = request.PageFrom; number <= lastParsed; number++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (text.IsFull)
                {
                    truncated = true;
                    break;
                }

                var pageText = ExtractPageText(document.GetPage(number)).Trim();
                returnedTo = number;
                if (pageText.Length == 0)
                {
                    pagesWithoutText++;
                    continue;
                }

                pagesWithText++;
                text.Append($"--- page {number} ---\n");
                text.Append(pageText);
                text.Append("\n\n");
            }

            if (pagesWithText == 0 && pagesWithoutText > 0) warnings.Add(ScannedPdfWarning);
            else if (pagesWithoutText > 0) warnings.Add($"{pagesWithoutText} page(s) have no text layer (scanned or drawing-only); OCR not supported.");
            return Result(file, pageCount, null, returnedTo is null ? null : request.PageFrom, returnedTo, null, text, truncated || text.Truncated, warnings);
        }
    }

    private static string ExtractPageText(UglyToad.PdfPig.Content.Page page)
    {
        try
        {
            return ContentOrderTextExtractor.GetText(page);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            var lines = page.Letters
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 0))
                .OrderByDescending(line => line.Key)
                .Select(line => string.Concat(line.OrderBy(letter => letter.StartBaseLine.X).Select(letter => letter.Value)));
            return string.Join('\n', lines);
        }
    }

    private AttachmentTextDto ReadWorkbook(string path, AttachmentScratchFileDto file, AttachmentTextRequest request, CancellationToken cancellationToken)
    {
        GuardPackage(path);
        var warnings = new List<string>();
        using var document = SpreadsheetDocument.Open(path, false);
        var workbook = document.WorkbookPart ?? throw new InvalidDataException("The workbook part is missing.");
        var sheets = workbook.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
        var names = sheets.Select(sheet => sheet.Name?.Value ?? string.Empty).ToList();
        var text = new BoundedText(request.MaximumCharacters);
        var returned = new List<string>();

        List<Sheet> selected;
        if (request.Sheet is not null)
        {
            selected = sheets.Where(sheet => string.Equals(sheet.Name?.Value, request.Sheet, StringComparison.OrdinalIgnoreCase)).Take(1).ToList();
            if (selected.Count == 0)
            {
                warnings.Add($"Sheet '{request.Sheet}' was not found; see sheet_names.");
                return Result(file, null, names, null, null, returned, text, false, warnings);
            }
        }
        else
        {
            selected = sheets.Take(_options.MaximumSheets).ToList();
        }

        var truncated = false;
        if (request.Sheet is null && sheets.Count > _options.MaximumSheets)
        {
            truncated = true;
            warnings.Add($"At most {_options.MaximumSheets} sheets are read per call; pass sheet to read a later one.");
        }

        var sharedStrings = workbook.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(item => item.InnerText).ToArray() ?? [];
        var cells = 0;
        var formulasWithoutValue = 0;
        foreach (var sheet in selected)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (text.IsFull || cells >= _options.MaximumCells)
            {
                truncated = true;
                break;
            }

            if (sheet.Id?.Value is not { } relationshipId || workbook.GetPartById(relationshipId) is not WorksheetPart worksheet) continue;
            returned.Add(sheet.Name?.Value ?? string.Empty);
            text.Append($"--- sheet {sheet.Name?.Value} ---\n");
            using var reader = OpenXmlReader.Create(worksheet);
            while (reader.Read())
            {
                if (reader.ElementType != typeof(Row) || !reader.IsStartElement) continue;
                cancellationToken.ThrowIfCancellationRequested();
                var row = (Row)reader.LoadCurrentElement()!;
                var values = new SortedDictionary<int, string>();
                var position = 0;
                foreach (var cell in row.Elements<Cell>())
                {
                    position = ColumnIndex(cell.CellReference?.Value) ?? position + 1;
                    var value = CellText(cell, sharedStrings);
                    if (value is null && cell.CellFormula is not null) formulasWithoutValue++;
                    if (string.IsNullOrWhiteSpace(value)) continue;
                    if (cells >= _options.MaximumCells) break;
                    values[position] = Flatten(value);
                    cells++;
                }

                if (values.Count > 0)
                {
                    var line = new StringBuilder();
                    for (var column = 1; column <= values.Keys.Last(); column++)
                    {
                        if (column > 1) line.Append('\t');
                        if (values.TryGetValue(column, out var value)) line.Append(value);
                    }

                    text.Append(line.Append('\n').ToString());
                }

                if (text.IsFull || cells >= _options.MaximumCells) break;
            }

            text.Append("\n");
        }

        if (cells >= _options.MaximumCells)
        {
            truncated = true;
            warnings.Add($"Stopped after {_options.MaximumCells} non-empty cells.");
        }

        if (formulasWithoutValue > 0) warnings.Add($"{formulasWithoutValue} formula cell(s) had no cached value and were left empty.");
        return Result(file, null, names, null, null, returned, text, truncated || text.Truncated, warnings);
    }

    private static string? CellText(Cell cell, string[] sharedStrings)
    {
        var type = cell.DataType?.Value;
        if (type == CellValues.InlineString) return cell.InlineString?.InnerText;
        var raw = cell.CellValue?.Text;
        if (raw is null) return null;
        if (type == CellValues.SharedString)
        {
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) && index >= 0 && index < sharedStrings.Length ? sharedStrings[index] : raw;
        }

        if (type == CellValues.Boolean) return raw == "1" ? "TRUE" : "FALSE";
        return raw;
    }

    private static int? ColumnIndex(string? reference)
    {
        if (string.IsNullOrEmpty(reference)) return null;
        var index = 0;
        foreach (var character in reference)
        {
            if (character is < 'A' or > 'Z') break;
            index = index * 26 + (character - 'A' + 1);
        }

        return index == 0 ? null : index;
    }

    private AttachmentTextDto ReadDocument(string path, AttachmentScratchFileDto file, AttachmentTextRequest request, CancellationToken cancellationToken)
    {
        GuardPackage(path);
        using var document = WordprocessingDocument.Open(path, false);
        var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidDataException("The document body is missing.");
        var text = new BoundedText(request.MaximumCharacters);
        foreach (var element in body.ChildElements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (text.IsFull) break;
            switch (element)
            {
                case Word.Paragraph paragraph:
                    AppendLine(text, ParagraphText(paragraph));
                    break;
                case Word.Table table:
                    foreach (var row in table.Elements<Word.TableRow>())
                    {
                        var cells = row.Elements<Word.TableCell>()
                            .Select(cell => Flatten(string.Join(' ', cell.Descendants<Word.Paragraph>().Select(ParagraphText).Where(value => value.Length > 0))))
                            .ToList();
                        if (cells.Any(value => value.Length > 0)) text.Append(string.Join('\t', cells) + "\n");
                    }

                    text.Append("\n");
                    break;
                default:
                    foreach (var paragraph in element.Descendants<Word.Paragraph>()) AppendLine(text, ParagraphText(paragraph));
                    break;
            }
        }

        return Result(file, null, null, null, null, null, text, text.Truncated, []);
    }

    private static void AppendLine(BoundedText text, string line)
    {
        if (line.Trim().Length > 0) text.Append(line + "\n");
    }

    private static string ParagraphText(Word.Paragraph paragraph)
    {
        var builder = new StringBuilder();
        foreach (var node in paragraph.Descendants())
        {
            switch (node)
            {
                case Word.Text value: builder.Append(value.Text); break;
                case Word.TabChar: builder.Append('\t'); break;
                case Word.Break or Word.CarriageReturn: builder.Append('\n'); break;
            }
        }

        return builder.ToString();
    }

    private AttachmentTextDto ReadPlainText(string path, AttachmentScratchFileDto file, AttachmentTextRequest request)
    {
        var bytes = File.ReadAllBytes(path);
        var warnings = new List<string>();
        string content;
        if (bytes is [0xEF, 0xBB, 0xBF, ..]) content = StrictUtf8.GetString(bytes, 3, bytes.Length - 3);
        else if (bytes is [0xFF, 0xFE, ..]) content = Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
        else if (bytes is [0xFE, 0xFF, ..]) content = Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        else
        {
            try { content = StrictUtf8.GetString(bytes); }
            catch (DecoderFallbackException)
            {
                content = Encoding.GetEncoding(1257).GetString(bytes);
                warnings.Add("Not valid UTF-8; decoded as Windows-1257 (Baltic).");
            }
        }

        var text = new BoundedText(request.MaximumCharacters);
        text.Append(content.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n'));
        return Result(file, null, null, null, null, null, text, text.Truncated, warnings);
    }

    /// <summary>Rejects OpenXML packages that would expand beyond the configured size or ratio before any XML is parsed.</summary>
    private void GuardPackage(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        long uncompressed = 0;
        long compressed = 0;
        foreach (var entry in archive.Entries)
        {
            uncompressed += entry.Length;
            compressed += entry.CompressedLength;
            if (uncompressed > _options.MaximumUncompressedBytes)
            {
                throw new OutlookMcpException(ErrorCodes.AttachmentTooLarge, "The document expands beyond the configured AttachmentText.MaximumUncompressedBytes limit.", "Open the file in its own application.");
            }
        }

        if (uncompressed > 0 && uncompressed > (long)_options.MaximumCompressionRatio * Math.Max(1, compressed))
        {
            throw new OutlookMcpException(ErrorCodes.AttachmentTooLarge, $"The document's compression ratio is above {_options.MaximumCompressionRatio}:1, which is typical of a zip bomb, so it was not parsed.", "Open the file in its own application if you trust it.");
        }
    }

    private static string Flatten(string value) => value.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ').Trim();

    private static AttachmentTextDto Result(AttachmentScratchFileDto file, int? pageCount, IReadOnlyList<string>? sheetNames, int? returnedFrom, int? returnedTo,
        IReadOnlyList<string>? returnedSheets, BoundedText text, bool truncated, IReadOnlyList<string> warnings)
    {
        var value = text.ToString().TrimEnd();
        return new(file.FileName, file.Extension, file.MimeType, file.SizeBytes, true, null, pageCount, sheetNames, returnedFrom, returnedTo, returnedSheets,
            value.Length, truncated, value, warnings);
    }

    private sealed class BoundedText(int maximum)
    {
        private readonly StringBuilder _builder = new();

        public bool Truncated { get; private set; }
        public bool IsFull => _builder.Length >= maximum;

        public void Append(string value)
        {
            var room = maximum - _builder.Length;
            if (room <= 0)
            {
                if (value.Length > 0) Truncated = true;
                return;
            }

            if (value.Length <= room)
            {
                _builder.Append(value);
                return;
            }

            _builder.Append(value, 0, room);
            Truncated = true;
        }

        public override string ToString() => _builder.ToString();
    }
}
