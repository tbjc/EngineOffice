using OfficeEngine.Tests.Support;
using System.Text;
using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Pdf;
using OfficeEngine.Pdf.LibreOffice;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Pdf;

/// <summary>
/// Integration tests of the LibreOffice headless adapter. They need the external engine and are
/// skipped, with an explicit reason, when it is not installed.
/// </summary>
public sealed class LibreOfficePdfConverterTests
{
    [LibreOfficeFact]
    public void Convert_WordDocument_ProducesAPdf()
    {
        using LibreOfficePdfConverter converter = new();
        using MemoryStream source = new();
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("Informe mensual de ventas");
            document.Save(source);
        }

        source.Position = 0;
        PdfConversionResult result = converter.Convert(source, PdfSourceFormat.WordDocument);

        Assert.True(result.Pdf.Length > 0);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(result.Pdf, 0, 5));
        Assert.True(result.Duration > TimeSpan.Zero);
    }

    [LibreOfficeFact]
    public void Convert_ExcelWorkbook_ProducesAPdf()
    {
        using LibreOfficePdfConverter converter = new();
        using MemoryStream source = new();
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0]["A1"].Value = "Importe";
            workbook.Worksheets[0]["A2"].Value = 1234.56;
            workbook.Save(source);
        }

        source.Position = 0;
        PdfConversionResult result = converter.Convert(source, PdfSourceFormat.ExcelWorkbook);

        Assert.True(result.Pdf.Length > 0);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(result.Pdf, 0, 5));
    }

    [LibreOfficeFact]
    public void Convert_UsesTheBytesItReceives()
    {
        using LibreOfficePdfConverter converter = new();
        byte[] first = CreateDocument("contenido primero");
        byte[] second = CreateDocument("contenido segundo, con otro texto");

        using MemoryStream firstStream = new(first);
        using MemoryStream secondStream = new(second);
        PdfConversionResult firstResult = converter.Convert(firstStream, PdfSourceFormat.WordDocument);
        PdfConversionResult secondResult = converter.Convert(secondStream, PdfSourceFormat.WordDocument);

        Assert.NotEqual(firstResult.Pdf.Length, secondResult.Pdf.Length);
    }

    [LibreOfficeFact]
    public void Convert_WithACorruptDocument_NeverReturnsGarbage()
    {
        using LibreOfficePdfConverter converter = new();
        using MemoryStream source = new(new byte[] { 1, 2, 3, 4, 5 });

        // The engine is tolerant: it may open the bytes as a text document and print it, or refuse
        // it. What it must never do is hand back something that is not a PDF.
        try
        {
            PdfConversionResult result = converter.Convert(source, PdfSourceFormat.WordDocument);
            Assert.Equal("%PDF-", Encoding.ASCII.GetString(result.Pdf, 0, 5));
        }
        catch (PdfConversionException)
        {
            // The engine refused the document, which is also a correct outcome.
        }
    }

    [LibreOfficeFact]
    public void Convert_WithATooShortTimeout_Fails()
    {
        using LibreOfficePdfConverter converter = new();
        using MemoryStream source = new();
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("contenido");
            document.Save(source);
        }

        source.Position = 0;
        PdfConversionOptions options = new() { Timeout = TimeSpan.FromMilliseconds(1) };

        PdfConversionException exception = Assert.Throws<PdfConversionException>(
            () => converter.Convert(source, PdfSourceFormat.WordDocument, options));
        Assert.Null(exception.ExitCode);
    }

    [LibreOfficeFact]
    public void Convert_LeavesNoTemporaryDirectoryBehind()
    {
        string root = Path.Combine(Path.GetTempPath(), "OfficeEngine.Pdf");
        int before = CountConversionDirectories(root);

        using LibreOfficePdfConverter converter = new();
        using MemoryStream source = new();
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("temporal");
            document.Save(source);
        }

        source.Position = 0;
        converter.Convert(source, PdfSourceFormat.WordDocument);

        Assert.Equal(before, CountConversionDirectories(root));
    }

    [Fact]
    public void Convert_WithAnEmptyStream_Throws()
    {
        using LibreOfficePdfConverter converter = new();
        using MemoryStream source = new();

        Assert.Throws<ArgumentException>(() => converter.Convert(source, PdfSourceFormat.WordDocument));
    }

    [Fact]
    public void Convert_WithANullStream_Throws()
    {
        using LibreOfficePdfConverter converter = new();

        Assert.Throws<ArgumentNullException>(() => converter.Convert(null!, PdfSourceFormat.WordDocument));
    }

    [Fact]
    public void Convert_WithAnUnknownFormat_Throws()
    {
        using LibreOfficePdfConverter converter = new();
        using MemoryStream source = new([1, 2, 3]);

        Assert.Throws<ArgumentOutOfRangeException>(() => converter.Convert(source, (PdfSourceFormat)99));
    }

    [Fact]
    public void Convert_WithAnInputLargerThanTheLimit_Throws()
    {
        using LibreOfficePdfConverter converter = new();
        using MemoryStream source = new(new byte[1024]);
        PdfConversionOptions options = new() { MaximumInputBytes = 16 };

        Assert.Throws<OfficeResourceLimitExceededException>(
            () => converter.Convert(source, PdfSourceFormat.WordDocument, options));
    }

    [Fact]
    public void Convert_WithAnInvalidTimeout_Throws()
    {
        using LibreOfficePdfConverter converter = new();
        using MemoryStream source = new([1, 2, 3]);
        PdfConversionOptions options = new() { Timeout = TimeSpan.Zero };

        Assert.Throws<ArgumentOutOfRangeException>(
            () => converter.Convert(source, PdfSourceFormat.WordDocument, options));
    }

    [Fact]
    public void Convert_WithAMissingEngine_Fails()
    {
        using LibreOfficePdfConverter converter = new(new LibreOfficePdfConverterOptions
        {
            ExecutablePath = Path.Combine(Path.GetTempPath(), "no-existe", "soffice.exe"),
        });
        using MemoryStream source = new(CreateDocument("contenido"));

        Assert.False(converter.IsAvailable);
        PdfConversionException exception = Assert.Throws<PdfConversionException>(
            () => converter.Convert(source, PdfSourceFormat.WordDocument));
        Assert.Null(exception.ExitCode);
    }

    [LibreOfficeFact]
    public void TheEngine_ReportsItsPathAndVersion()
    {
        using LibreOfficePdfConverter converter = new();

        Assert.True(converter.IsAvailable);
        Assert.NotNull(converter.ExecutablePath);
        Assert.NotNull(converter.Version);
        Assert.True(char.IsDigit(converter.Version![0]));
    }

    private static byte[] CreateDocument(string text)
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph(text);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static int CountConversionDirectories(string root)
        => Directory.Exists(root) ? Directory.GetDirectories(root).Length : 0;
}
