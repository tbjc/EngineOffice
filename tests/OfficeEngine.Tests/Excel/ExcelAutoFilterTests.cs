using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>The auto filter of a worksheet: the range it covers and its removal.</summary>
public sealed class ExcelAutoFilterTests
{
    [Fact]
    public void AutoFilter_Set_WritesTheRange()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.AutoFilter = new ExcelAutoFilter("A1:D10");

            Assert.Equal("A1:D10", sheet.AutoFilter!.Range);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.Equal("A1:D10", OoxmlValidation.ReadXlsxAutoFilter(bytes, "Sheet1"));

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Assert.Equal("A1:D10", reopened.Worksheets[0].AutoFilter!.Range);
    }

    [Fact]
    public void AutoFilter_SetNull_RemovesIt()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.AutoFilter = new ExcelAutoFilter("A1:D10");

        Assert.True(sheet.AutoFilter is not null);
        sheet.AutoFilter = null;

        Assert.Null(sheet.AutoFilter);
        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("autoFilter", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void AutoFilter_ReadsTheFilterOfAnotherWorkbook()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithAutoFilter("A1:D10");

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        Assert.Equal("A1:D10", workbook.Worksheets[0].AutoFilter!.Range);
    }

    [Fact]
    public void AutoFilter_WithAnInvalidRange_Throws()
    {
        Assert.Throws<ArgumentException>(() => new ExcelAutoFilter("B2:A1"));
        Assert.Throws<ArgumentException>(() => new ExcelAutoFilter("Hoja1!A1"));
    }

    [Fact]
    public void OpenedWorkbook_WithAnAutoFilter_DoesNotReportItAsUnmodelledContent()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithAutoFilter("A1:D10");

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        Assert.Empty(workbook.Diagnostics);
    }

    [Fact]
    public void DisposedWorkbook_AutoFilterThrowsObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].AutoFilter);
    }
}
