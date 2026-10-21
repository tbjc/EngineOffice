using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>The print settings of a sheet: margins, page setup and print options.</summary>
public sealed class ExcelPrintSettingsTests
{
    [Fact]
    public void PageMargins_Set_WritesTheMargins()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageMargins = new ExcelPageMargins(1, 1.25, 0.9, 0.8, 0.2, 0.4);

            Assert.Equal(1, sheet.PageMargins!.Left);
            Assert.Equal(0.4, sheet.PageMargins.Footer);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        (double Left, double Right, double Top, double Bottom, double Header, double Footer)? margins =
            OoxmlValidation.ReadXlsxPageMargins(bytes, "Sheet1");
        Assert.NotNull(margins);
        Assert.Equal(1.25, margins!.Value.Right);
        Assert.Equal(0.9, margins.Value.Top);

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Assert.Equal(0.8, reopened.Worksheets[0].PageMargins!.Bottom);
    }

    [Fact]
    public void PageMargins_Null_RemovesThem()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.PageMargins = new ExcelPageMargins(1, 1, 1, 1);

        sheet.PageMargins = null;

        Assert.Null(sheet.PageMargins);
        Assert.Null(OoxmlValidation.ReadXlsxPageMargins(SaveBytes.ToBytes(workbook), "Sheet1"));
    }

    [Fact]
    public void PageMargins_WithInvalidValues_Throws()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelPageMargins(-1, 1, 1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelPageMargins(1, 1, 1, 200));
    }

    [Fact]
    public void PageSetup_Set_WritesTheSetup()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup = new ExcelPageSetup(
                ExcelPageOrientation.Landscape,
                ExcelPaperSize.A3,
                scale: 150);

            Assert.Equal(ExcelPageOrientation.Landscape, sheet.PageSetup!.Orientation);
            Assert.Equal(ExcelPaperSize.A3, sheet.PageSetup.PaperSize);
            Assert.Equal(150, sheet.PageSetup.Scale);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        (uint PaperSize, string Orientation, uint? Scale, uint? FitToWidth, uint? FitToHeight)? setup =
            OoxmlValidation.ReadXlsxPageSetup(bytes, "Sheet1");
        Assert.NotNull(setup);
        Assert.Equal(8U, setup!.Value.PaperSize);
        Assert.Equal("landscape", setup.Value.Orientation);
        Assert.Equal(150U, setup.Value.Scale);
    }

    [Fact]
    public void PageSetup_WithFitToPage_AsksTheSheetToFit()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup = new ExcelPageSetup(fitToWidth: 1, fitToHeight: 0);

            Assert.Equal(1, sheet.PageSetup!.FitToWidth);
            Assert.Equal(0, sheet.PageSetup.FitToHeight);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        string sheetXml = TestPackage.ReadEntryText(bytes, "xl/worksheets/sheet1.xml");
        Assert.Contains("fitToPage=\"1\"", sheetXml, StringComparison.Ordinal);
        Assert.Contains("fitToWidth=\"1\"", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void PageSetup_Null_RemovesIt()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.PageSetup = new ExcelPageSetup();

        sheet.PageSetup = null;

        Assert.Null(sheet.PageSetup);
        Assert.Null(OoxmlValidation.ReadXlsxPageSetup(SaveBytes.ToBytes(workbook), "Sheet1"));
    }

    [Fact]
    public void PageSetup_WithInvalidValues_Throws()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelPageSetup(scale: 5));
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelPageSetup(scale: 500));
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelPageSetup(fitToWidth: -1));
        Assert.Throws<ArgumentException>(() => new ExcelPageSetup(scale: 100, fitToWidth: 1));
    }

    [Fact]
    public void PrintOptions_Set_WritesTheOptions()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PrintOptions = new ExcelPrintOptions(
                horizontalCentered: true,
                verticalCentered: true,
                gridLines: true,
                headings: true);

            Assert.True(sheet.PrintOptions!.HorizontalCentered);
            Assert.True(sheet.PrintOptions.GridLines);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        (bool HorizontalCentered, bool VerticalCentered, bool GridLines, bool Headings)? options =
            OoxmlValidation.ReadXlsxPrintOptions(bytes, "Sheet1");
        Assert.NotNull(options);
        Assert.True(options!.Value.HorizontalCentered);
        Assert.True(options.Value.Headings);
    }

    [Fact]
    public void PrintOptions_Null_RemovesThem()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.PrintOptions = new ExcelPrintOptions(gridLines: true);

        sheet.PrintOptions = null;

        Assert.Null(sheet.PrintOptions);
        Assert.Null(OoxmlValidation.ReadXlsxPrintOptions(SaveBytes.ToBytes(workbook), "Sheet1"));
    }

    [Fact]
    public void PageSetup_ReadsTheSetupOfAnotherWorkbook()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithPrintSettings(
            left: 1.5,
            orientation: "landscape",
            paperSize: 5U,
            scale: 90U,
            horizontalCentered: true);

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        ExcelPageMargins margins = workbook.Worksheets[0].PageMargins!;
        ExcelPageSetup setup = workbook.Worksheets[0].PageSetup!;
        Assert.Equal(1.5, margins.Left);
        Assert.Equal(ExcelPageOrientation.Landscape, setup.Orientation);
        Assert.Equal(ExcelPaperSize.Legal, setup.PaperSize);
        Assert.Equal(90, setup.Scale);
        Assert.True(workbook.Worksheets[0].PrintOptions!.HorizontalCentered);
    }

    [Fact]
    public void DisposedWorkbook_PrintSettingsThrowObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].PageMargins);
        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].PageSetup);
        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].PrintOptions);
    }
}
