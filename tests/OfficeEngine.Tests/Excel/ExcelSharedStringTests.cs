using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>
/// The shared string table: the same text is stored once and every cell that repeats it points to
/// the same entry.
/// </summary>
public sealed class ExcelSharedStringTests
{
    [Fact]
    public void Strings_AreWrittenToTheSharedStringTable()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet["A1"].Value = "Repetido";
            sheet["A2"].Value = "Repetido";
            sheet["B1"].Value = "Otro";

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        (IReadOnlyList<string> Items, uint? Count, uint? UniqueCount) table =
            OoxmlValidation.ReadXlsxSharedStrings(bytes);
        Assert.Equal(["Repetido", "Otro"], table.Items);
        Assert.Equal(2U, table.UniqueCount);
        Assert.Equal(3U, table.Count);

        string sheetXml = TestPackage.ReadEntryText(bytes, "xl/worksheets/sheet1.xml");
        Assert.Contains("t=\"s\"", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void Strings_RoundTripThroughSaveAndOpen()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0]["A1"].Value = "Con espacios ";
            workbook.Worksheets[0]["A2"].Value = "Otro";

            bytes = SaveBytes.ToBytes(workbook);
        }

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Assert.Equal("Con espacios ", reopened.Worksheets[0]["A1"].Value);
        Assert.Equal("Otro", reopened.Worksheets[0]["A2"].Value);
    }

    [Fact]
    public void ReplacingAString_KeepsTheTableConsistent()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet["A1"].Value = "Primero";
            sheet["A1"].Value = "Segundo";

            Assert.Equal("Segundo", sheet["A1"].Value);

            bytes = SaveBytes.ToBytes(workbook);
        }

        (IReadOnlyList<string> Items, uint? Count, uint? UniqueCount) table =
            OoxmlValidation.ReadXlsxSharedStrings(bytes);
        Assert.Equal(["Primero", "Segundo"], table.Items);
        Assert.Equal(1U, table.Count);
    }

    [Fact]
    public void ClearingAString_RemovesItsReference()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet["A1"].Value = "Temporal";
            sheet["A1"].Value = null;

            Assert.Null(sheet["A1"].Value);

            bytes = SaveBytes.ToBytes(workbook);
        }

        (IReadOnlyList<string> Items, uint? Count, uint? UniqueCount) table =
            OoxmlValidation.ReadXlsxSharedStrings(bytes);
        Assert.Equal(0U, table.Count);
    }

    [Fact]
    public void AnEmptyString_IsStored()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0]["A1"].Value = string.Empty;

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Assert.Equal(string.Empty, reopened.Worksheets[0]["A1"].Value);
    }

    [Fact]
    public void TheSharedStringsOfAnotherWorkbook_AreStillRead()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithSharedStringsAndMergedCells("Hola", "Mundo");

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        Assert.Equal("Hola", workbook.Worksheets[0]["A1"].Value);
        Assert.Equal("Mundo", workbook.Worksheets[0]["B1"].Value);
    }
}
