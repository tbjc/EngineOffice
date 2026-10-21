using OfficeEngine.Excel;
using OfficeEngine.Legacy;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Legacy;

/// <summary>
/// The conversion of the legacy formats: a conversion, not a reader, and a result that opens with
/// the modern API.
/// </summary>
public sealed class LegacyConverterTests
{
    [LibreOfficeFact]
    public void Convert_WordDocument_ProducesAModernDocument()
    {
        byte[] docx;
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("Informe heredado");
            document.AddParagraph("Segunda línea del informe");
            using MemoryStream stream = new();
            document.Save(stream);
            docx = stream.ToArray();
        }

        byte[] doc = LegacyFixtures.ToLegacyWord(docx);
        LegacyConverter converter = new();
        using MemoryStream source = new(doc);
        LegacyConversionResult result = converter.Convert(source, LegacyDocumentFormat.Word97To2003);

        Assert.Equal(LegacyDocumentFormat.Word97To2003, result.Format);
        Assert.True(result.Document.Length > 0);
        Assert.True(result.Duration >= TimeSpan.Zero);

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(result.Document));
        Assert.Contains("Informe heredado", reopened.Text, StringComparison.Ordinal);
    }

    [LibreOfficeFact]
    public void Convert_ExcelWorkbook_ProducesAModernWorkbook()
    {
        byte[] xlsx;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0]["A1"].Value = "Total";
            workbook.Worksheets[0]["A2"].Value = 1234.5d;
            using MemoryStream stream = new();
            workbook.Save(stream);
            xlsx = stream.ToArray();
        }

        byte[] xls = LegacyFixtures.ToLegacyExcel(xlsx);
        LegacyConverter converter = new();
        using MemoryStream source = new(xls);
        LegacyConversionResult result = converter.Convert(source, LegacyDocumentFormat.Excel97To2003);

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(result.Document));
        Assert.Equal("Total", reopened.Worksheets[0]["A1"].Value);
        Assert.Equal(1234.5d, reopened.Worksheets[0]["A2"].Value);
    }

    [LibreOfficeFact]
    public void Convert_WithACorruptDocument_NeverReturnsSomethingBroken()
    {
        LegacyConverter converter = new();
        using MemoryStream source = new(new byte[] { 1, 2, 3, 4, 5 });

        // The engine is tolerant: it may open the bytes as a text document. What it must never do
        // is return something that is not a document.
        try
        {
            LegacyConversionResult result = converter.Convert(source, LegacyDocumentFormat.Word97To2003);
            using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(result.Document));
            Assert.NotNull(reopened);
        }
        catch (LegacyConversionException)
        {
            // The engine refused the document, which is also a correct outcome.
        }
    }

    [Fact]
    public void Convert_WithAMissingEngine_Fails()
    {
        LegacyConverter converter = new(new LegacyConverterOptions
        {
            ExecutablePath = Path.Combine(Path.GetTempPath(), "no-existe", "soffice.exe"),
        });
        using MemoryStream source = new(new byte[] { 1, 2, 3 });

        Assert.False(converter.IsAvailable);
        Assert.Throws<LegacyConversionException>(
            () => converter.Convert(source, LegacyDocumentFormat.Word97To2003));
    }

    [Fact]
    public void Convert_WithAnEmptyStream_Throws()
    {
        LegacyConverter converter = new();
        using MemoryStream source = new();

        Assert.Throws<ArgumentException>(
            () => converter.Convert(source, LegacyDocumentFormat.Word97To2003));
    }

    [Fact]
    public void Convert_WithANullStream_Throws()
    {
        LegacyConverter converter = new();

        Assert.Throws<ArgumentNullException>(
            () => converter.Convert(null!, LegacyDocumentFormat.Word97To2003));
    }

    [Fact]
    public void Convert_WithAnUnknownFormat_Throws()
    {
        LegacyConverter converter = new();
        using MemoryStream source = new(new byte[] { 1, 2, 3 });

        Assert.Throws<ArgumentOutOfRangeException>(
            () => converter.Convert(source, (LegacyDocumentFormat)99));
    }
}
