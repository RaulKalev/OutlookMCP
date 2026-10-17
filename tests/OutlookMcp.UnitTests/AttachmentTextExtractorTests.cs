using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Errors;
using OutlookMcp.Contracts;
using OutlookMcp.Infrastructure.Attachments;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Writer;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace OutlookMcp.UnitTests;

public sealed class AttachmentTextExtractorTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "outlook-mcp-attachment-tests", Guid.NewGuid().ToString("N"));

    public AttachmentTextExtractorTests()
    {
        Directory.CreateDirectory(_directory);
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    [Fact]
    public void Pdf_ReturnsRequestedPageRangeInOrder()
    {
        var path = WritePdf("Esimene leht", "Teine leht", "Kolmas leht");
        var result = Extract(path, new AttachmentTextRequest(20_000, 2, 3, null));

        Assert.True(result.Supported);
        Assert.Equal(3, result.PageCount);
        Assert.Equal(2, result.ReturnedFrom);
        Assert.Equal(3, result.ReturnedTo);
        Assert.False(result.Truncated);
        Assert.DoesNotContain("Esimene", result.Text, StringComparison.Ordinal);
        Assert.True(result.Text.IndexOf("--- page 2 ---", StringComparison.Ordinal) < result.Text.IndexOf("--- page 3 ---", StringComparison.Ordinal));
        Assert.Contains("Kolmas leht", result.Text, StringComparison.Ordinal);
        Assert.Equal(result.Text.Length, result.CharactersReturned);
        Assert.True(result.UntrustedContent);
    }

    [Fact]
    public void Pdf_TruncatesAtMaximumCharacters()
    {
        var path = WritePdf("Esimene leht", "Teine leht", "Kolmas leht");
        var result = Extract(path, new AttachmentTextRequest(30, 1, null, null));

        Assert.True(result.Truncated);
        Assert.Equal(30, result.CharactersReturned);
        Assert.Equal(1, result.ReturnedFrom);
        Assert.True(result.ReturnedTo < 3);
    }

    [Fact]
    public void Pdf_PageFromAfterLastPageReturnsWarning()
    {
        var path = WritePdf("Ainus leht");
        var result = Extract(path, new AttachmentTextRequest(20_000, 5, null, null));

        Assert.Equal(1, result.PageCount);
        Assert.Empty(result.Text);
        Assert.Contains(result.Warnings, warning => warning.Contains("after the last page", StringComparison.Ordinal));
    }

    [Fact(Skip = "PdfPig's PdfDocumentBuilder cannot write encrypted PDFs; the PdfDocumentEncryptedException path is covered by the manual mailbox check instead.")]
    public void Pdf_EncryptedDocumentReturnsWarning()
    {
    }

    [Fact]
    public void Docx_ReturnsParagraphsAndTableRows()
    {
        var path = Path.Combine(_directory, "seletuskiri.docx");
        using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Word.Document(new Word.Body(
                new Word.Paragraph(new Word.Run(new Word.Text("Nõrkvoolu seletuskiri"))),
                new Word.Table(
                    new Word.TableRow(Cell("Tähis"), Cell("Kirjeldus")),
                    new Word.TableRow(Cell("ATS-1"), Cell("Tulekahjusignalisatsioon")))));
        }

        var result = Extract(path, new AttachmentTextRequest(20_000, 1, null, null));

        Assert.True(result.Supported);
        Assert.Contains("Nõrkvoolu seletuskiri", result.Text, StringComparison.Ordinal);
        Assert.Contains("Tähis\tKirjeldus\nATS-1\tTulekahjusignalisatsioon", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void Xlsx_ReadsAllSheetsSkipsEmptyRowsAndUsesCachedFormulaValue()
    {
        var path = WriteWorkbook();
        var result = Extract(path, new AttachmentTextRequest(20_000, 1, null, null));

        Assert.Equal(["Joonised", "Muudatused"], result.SheetNames);
        Assert.Equal(["Joonised", "Muudatused"], result.ReturnedSheets);
        Assert.Contains("Nr\tNimetus\n1\tEN-1\t2", result.Text, StringComparison.Ordinal);
        Assert.Contains("Rev\tB", result.Text, StringComparison.Ordinal);
        Assert.False(result.Truncated);
    }

    [Fact]
    public void Xlsx_SheetFilterReturnsOnlyThatSheet()
    {
        var path = WriteWorkbook();
        var result = Extract(path, new AttachmentTextRequest(20_000, 1, null, "muudatused"));

        Assert.Equal(["Muudatused"], result.ReturnedSheets);
        Assert.Contains("Rev\tB", result.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("EN-1", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void Xlsx_StopsAtCellLimit()
    {
        var path = WriteWorkbook();
        var result = Extract(path, new AttachmentTextRequest(20_000, 1, null, null), options => options.MaximumCells = 3);

        Assert.True(result.Truncated);
        Assert.DoesNotContain("EN-1", result.Text, StringComparison.Ordinal);
        Assert.Contains(result.Warnings, warning => warning.Contains("3 non-empty cells", StringComparison.Ordinal));
    }

    [Fact]
    public void Csv_DetectsWindows1257()
    {
        const string content = "Tähis;Õpetaja;šahtid;žürii;öö;ülevaade";
        var path = Path.Combine(_directory, "tabel.csv");
        File.WriteAllBytes(path, Encoding.GetEncoding(1257).GetBytes(content));

        var result = Extract(path, new AttachmentTextRequest(20_000, 1, null, null));

        Assert.Equal(content, result.Text);
        Assert.Contains(result.Warnings, warning => warning.Contains("Windows-1257", StringComparison.Ordinal));
    }

    [Fact]
    public void Txt_ReadsUtf8WithBom()
    {
        const string content = "Protokoll: õ ä ö ü š ž";
        var path = Path.Combine(_directory, "protokoll.txt");
        File.WriteAllText(path, content, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

        var result = Extract(path, new AttachmentTextRequest(20_000, 1, null, null));

        Assert.Equal(content, result.Text);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void UnsupportedExtensionReturnsMetadataOnly()
    {
        var path = Path.Combine(_directory, "plaan.dwg");
        File.WriteAllBytes(path, [1, 2, 3]);

        var result = Extract(path, new AttachmentTextRequest(20_000, 1, null, null));

        Assert.False(result.Supported);
        Assert.Contains("CAD", result.UnsupportedReason, StringComparison.Ordinal);
        Assert.Empty(result.Text);
        Assert.Null(new AttachmentTextExtractor(new OutlookMcpOptions()).GetUnsupportedReason(".pdf"));
    }

    [Fact]
    public void OversizedFileIsRefused()
    {
        var path = Path.Combine(_directory, "suur.txt");
        File.WriteAllBytes(path, new byte[4_096]);

        var error = Assert.Throws<OutlookMcpException>(() => Extract(path, new AttachmentTextRequest(20_000, 1, null, null), options => options.MaximumFileBytes = 1_024));

        Assert.Equal(ErrorCodes.AttachmentTooLarge, error.Code);
    }

    [Fact]
    public void ZipBombRatioIsRefusedBeforeParsing()
    {
        var path = Path.Combine(_directory, "pomm.docx");
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            using var stream = archive.CreateEntry("word/document.xml", CompressionLevel.SmallestSize).Open();
            stream.Write(new byte[5 * 1024 * 1024]);
        }

        var error = Assert.Throws<OutlookMcpException>(() => Extract(path, new AttachmentTextRequest(20_000, 1, null, null)));

        Assert.Equal(ErrorCodes.AttachmentTooLarge, error.Code);
        Assert.Contains("compression ratio", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DamagedFileReportsParseFailure()
    {
        var path = Path.Combine(_directory, "katki.xlsx");
        File.WriteAllBytes(path, Encoding.ASCII.GetBytes("not a zip package"));

        var error = Assert.Throws<OutlookMcpException>(() => Extract(path, new AttachmentTextRequest(20_000, 1, null, null)));

        Assert.Equal(ErrorCodes.AttachmentParseFailed, error.Code);
    }

    private AttachmentTextDto Extract(string path, AttachmentTextRequest request, Action<AttachmentTextOptions>? configure = null)
    {
        var options = new OutlookMcpOptions();
        configure?.Invoke(options.AttachmentText);
        var file = new AttachmentScratchFileDto(Path.GetFileName(path), Path.GetExtension(path), null, new FileInfo(path).Length, path, null);
        return new AttachmentTextExtractor(options).Extract(path, file, request, CancellationToken.None);
    }

    private string WritePdf(params string[] pages)
    {
        var builder = new PdfDocumentBuilder();
        var font = builder.AddStandard14Font(Standard14Font.Helvetica);
        foreach (var text in pages) builder.AddPage(PageSize.A4).AddText(text, 12, new PdfPoint(50, 700), font);
        var path = Path.Combine(_directory, "joonised.pdf");
        File.WriteAllBytes(path, builder.Build());
        return path;
    }

    private string WriteWorkbook()
    {
        var path = Path.Combine(_directory, "joonised.xlsx");
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbook = document.AddWorkbookPart();
        workbook.Workbook = new Workbook();
        var sheets = workbook.Workbook.AppendChild(new Sheets());
        AddSheet(workbook, sheets, 1, "Joonised",
            new Row(Text("A1", "Nr"), Text("B1", "Nimetus")) { RowIndex = 1 },
            new Row(new Cell { CellReference = "A2" }) { RowIndex = 2 },
            new Row(
                new Cell { CellReference = "A3", CellValue = new CellValue("1") },
                Text("B3", "EN-1"),
                new Cell { CellReference = "C3", CellFormula = new CellFormula("A3*2"), CellValue = new CellValue("2") }) { RowIndex = 3 });
        AddSheet(workbook, sheets, 2, "Muudatused", new Row(Text("A1", "Rev"), Text("B1", "B")) { RowIndex = 1 });
        return path;
    }

    private static void AddSheet(WorkbookPart workbook, Sheets sheets, uint id, string name, params Row[] rows)
    {
        var worksheet = workbook.AddNewPart<WorksheetPart>();
        worksheet.Worksheet = new Worksheet(new SheetData(rows.Cast<OpenXmlElement>()));
        sheets.Append(new Sheet { Id = workbook.GetIdOfPart(worksheet), SheetId = id, Name = name });
    }

    private static Cell Text(string reference, string value) =>
        new() { CellReference = reference, DataType = CellValues.InlineString, InlineString = new InlineString(new Text(value)) };

    private static Word.TableCell Cell(string value) => new(new Word.Paragraph(new Word.Run(new Word.Text(value))));

    public void Dispose()
    {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
    }
}
