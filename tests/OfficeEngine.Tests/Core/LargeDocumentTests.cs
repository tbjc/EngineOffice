using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Core;

/// <summary>
/// The limits of the in-memory model, measured: a document that is large but reasonable must open,
/// save and reopen without surprises.
/// </summary>
public sealed class LargeDocumentTests
{
    [Fact]
    public void ALargeWordDocument_RoundTrips()
    {
        const int paragraphs = 10000;
        byte[] package;

        using (WordDocument document = WordDocument.Create())
        {
            for (int index = 0; index < paragraphs; index++)
            {
                document.AddParagraph($"Párrafo {index} del documento grande");
            }

            using MemoryStream stream = new();
            document.Save(stream);
            package = stream.ToArray();
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(package));
        Assert.Equal(paragraphs, reopened.Paragraphs.Count);
        Assert.Equal("Párrafo 9999 del documento grande", reopened.Paragraphs[paragraphs - 1].Text);
    }

    [Fact]
    public void ALargeExcelWorkbook_RoundTrips()
    {
        const int rows = 10000;
        byte[] package;

        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 1; row <= rows; row++)
            {
                sheet[row, 1].Value = $"Fila {row}";
                sheet[row, 2].Value = row;
            }

            using MemoryStream stream = new();
            workbook.Save(stream);
            package = stream.ToArray();
        }

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(package));
        Worksheet reopenedSheet = reopened.Worksheets[0];
        Assert.Equal("Fila 10000", reopenedSheet[rows, 1].Value);
        Assert.Equal(10000d, reopenedSheet[rows, 2].Value);
    }

    [Fact]
    public void ADocumentThatExpandsBeyondTheLimit_IsRejected()
    {
        byte[] package;
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph(new string('a', 4096));
            using MemoryStream stream = new();
            document.Save(stream);
            package = stream.ToArray();
        }

        OfficeEngineLimits limits = new() { MaxTotalUncompressedSizeBytes = 1024 };
        Assert.Throws<OfficeResourceLimitExceededException>(
            () => WordDocument.Open(TestPackage.ToStream(package), leaveOpen: false, limits));
    }
}
