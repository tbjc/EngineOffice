using System.Text;
using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Security;

/// <summary>
/// The loading of hostile packages: deep nesting, document type definitions, external entities and
/// malformed XML must fail with a typed error, and never take the process down.
/// </summary>
public sealed class HostilePackageTests
{
    [Fact]
    public void Word_WithADeeplyNestedPart_ThrowsATypedError()
    {
        byte[] package = CreateWordDocument();
        byte[] hostile = TestPackage.WithEntryText(
            package,
            "word/document.xml",
            DeepXml(5000, "w:document", "w:p", WordNamespace));

        // The limit on the compression ratio is raised so that the check under test is the depth.
        OfficeEngineLimits limits = new() { MaxEntryCompressionRatio = 10000 };
        OfficeResourceLimitExceededException error = Assert.Throws<OfficeResourceLimitExceededException>(
            () => WordDocument.Open(TestPackage.ToStream(hostile), leaveOpen: false, limits));
        Assert.Contains("deeper than the configured maximum", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Excel_WithADeeplyNestedPart_ThrowsATypedError()
    {
        byte[] package = CreateExcelWorkbook();
        byte[] hostile = TestPackage.WithEntryText(package, "xl/worksheets/sheet1.xml", DeepXml(5000, "worksheet", "row"));

        OfficeEngineLimits limits = new() { MaxEntryCompressionRatio = 10000 };
        Assert.Throws<OfficeResourceLimitExceededException>(
            () => ExcelWorkbook.Open(TestPackage.ToStream(hostile), leaveOpen: false, limits));
    }

    [Fact]
    public void Word_WithADocumentTypeDefinition_ThrowsATypedError()
    {
        byte[] package = CreateWordDocument();
        string dtd = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <!DOCTYPE w:document [ <!ELEMENT w:document ANY> ]>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>
            """;
        byte[] hostile = TestPackage.WithEntryText(package, "word/document.xml", dtd);

        Assert.Throws<InvalidOfficeFileException>(() => WordDocument.Open(TestPackage.ToStream(hostile)));
    }

    [Fact]
    public void Excel_WithAnExternalEntity_ThrowsATypedError()
    {
        byte[] package = CreateExcelWorkbook();
        string xxe = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <!DOCTYPE worksheet [ <!ENTITY xxe SYSTEM "file:///C:/Windows/win.ini"> ]>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>
            """;
        byte[] hostile = TestPackage.WithEntryText(package, "xl/worksheets/sheet1.xml", xxe);

        InvalidOfficeFileException error = Assert.Throws<InvalidOfficeFileException>(
            () => ExcelWorkbook.Open(TestPackage.ToStream(hostile)));
        Assert.Contains("xl/worksheets/sheet1.xml", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Excel_WithAnInvalidCharacterInXml_ThrowsATypedError()
    {
        byte[] package = CreateExcelWorkbook();
        string invalid = "<?xml version=\"1.0\"?><worksheet><sheetData><row><c><v></v></c></row></sheetData></worksheet>";
        byte[] hostile = TestPackage.WithEntryText(package, "xl/worksheets/sheet1.xml", invalid);

        Assert.Throws<InvalidOfficeFileException>(() => ExcelWorkbook.Open(TestPackage.ToStream(hostile)));
    }

    [Fact]
    public void TheXmlDepthLimit_IsConfigurable()
    {
        byte[] package = CreateExcelWorkbook();
        byte[] hostile = TestPackage.WithEntryText(
            package,
            "xl/worksheets/sheet1.xml",
            DeepXml(20, "worksheet", "row", SpreadsheetNamespace));

        OfficeEngineLimits strict = new() { MaximumXmlDepth = 10 };
        Assert.Throws<OfficeResourceLimitExceededException>(
            () => ExcelWorkbook.Open(TestPackage.ToStream(hostile), leaveOpen: false, strict));

        // With a limit that allows the depth, the guard lets the part through.
        OfficeEngineLimits relaxed = new() { MaximumXmlDepth = 1000 };
        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(hostile), leaveOpen: false, relaxed);
        Assert.NotNull(workbook.Worksheets);
    }

    [Fact]
    public void AValidDocument_OpensWithTheGuardEnabled()
    {
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(CreateWordDocument()));
        Assert.Single(document.Paragraphs);

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(CreateExcelWorkbook()));
        Assert.Single(workbook.Worksheets);
    }

    private static byte[] CreateWordDocument()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("Texto");
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static byte[] CreateExcelWorkbook()
    {
        using MemoryStream stream = new();
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0]["A1"].Value = "Texto";
            workbook.Save(stream);
        }

        return stream.ToArray();
    }

    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static string DeepXml(int depth, string root, string element, string? xmlNamespace = null)
    {
        StringBuilder builder = new();
        builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        builder.Append('<').Append(root);
        if (xmlNamespace is not null)
        {
            int colon = root.IndexOf(':');
            if (colon > 0)
            {
                builder.Append(" xmlns:").Append(root[..colon]).Append('=');
            }
            else
            {
                builder.Append(" xmlns=");
            }

            builder.Append('"').Append(xmlNamespace).Append('"');
        }

        builder.Append('>');
        for (int index = 0; index < depth; index++)
        {
            builder.Append('<').Append(element).Append('>');
        }

        for (int index = 0; index < depth; index++)
        {
            builder.Append("</").Append(element).Append('>');
        }

        builder.Append("</").Append(root).Append('>');
        return builder.ToString();
    }
}
