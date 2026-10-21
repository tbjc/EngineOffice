using DocumentFormat.OpenXml.Validation;
using OfficeEngine;
using OfficeEngine.Diagnostics;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Tests.Excel;

/// <summary>
/// Excel interoperability checks: produced packages are verified with the Open XML SDK,
/// and SDK-produced fixtures are read back through OfficeEngine.
/// </summary>
public sealed class ExcelInteropTests
{
    [Fact]
    public void ProducedPackage_IsAValidSpreadsheetDocument()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet["A1"].Value = "x";
        sheet["B1"].Value = 3.5;
        sheet["C1"].Value = false;
        _ = workbook.Worksheets.Add("Segunda");

        IReadOnlyList<ValidationErrorInfo> errors = OoxmlValidation.ValidateXlsx(SaveBytes.ToBytes(workbook));

        Assert.Empty(errors);
    }

    [Fact]
    public void ProducedPackage_ContainsSharedStringsAndIsReadableByTheSdk()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Worksheets[0]["A1"].Value = "Contenido";
        byte[] bytes = SaveBytes.ToBytes(workbook);

        string sheetXml = TestPackage.ReadEntryText(bytes, "xl/worksheets/sheet1.xml");
        Assert.Contains("t=\"s\"", sheetXml, StringComparison.Ordinal);
        Assert.True(TestPackage.HasEntry(bytes, "xl/sharedStrings.xml"));
        Assert.Contains("Contenido", TestPackage.ReadEntryText(bytes, "xl/sharedStrings.xml"), StringComparison.Ordinal);

        Assert.Equal("Contenido", OoxmlValidation.ReadXlsxCellText(bytes, "Sheet1", "A1"));
    }

    [Fact]
    public void ProducedPackage_WithCellWrittenByIndex_IsValidAndReadableByTheSdk()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Worksheets[0][2, 2].Value = "por índice";

        byte[] bytes = SaveBytes.ToBytes(workbook);

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.Equal("por índice", OoxmlValidation.ReadXlsxCellText(bytes, "Sheet1", "B2"));
    }

    [Fact]
    public void ProducedPackage_WithRangeValues_IsValidAndReadableByTheSdk()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Worksheets[0].GetRange("A1:B2").SetValues(new object?[,]
        {
            { "a", 1.0 },
            { true, null },
        });

        byte[] bytes = SaveBytes.ToBytes(workbook);

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.Equal("a", OoxmlValidation.ReadXlsxCellText(bytes, "Sheet1", "A1"));
        Assert.Equal("1", OoxmlValidation.ReadXlsxCellText(bytes, "Sheet1", "B1"));
        Assert.Equal("1", OoxmlValidation.ReadXlsxCellText(bytes, "Sheet1", "A2"));
        Assert.Null(OoxmlValidation.ReadXlsxCellText(bytes, "Sheet1", "B2"));
    }

    [Fact]
    public void ProducedPackage_WithDateAndNumberFormats_IsValidAndReadableByTheSdk()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet["A1"].SetDateTime(new DateTime(2024, 1, 31), "yyyy-mm-dd");
        sheet["B1"].Value = 1234.5;
        sheet["B1"].NumberFormat = "#,##0.00";

        byte[] bytes = SaveBytes.ToBytes(workbook);

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.Equal("45322", OoxmlValidation.ReadXlsxCellText(bytes, "Sheet1", "A1"));
        Assert.Equal("1234.5", OoxmlValidation.ReadXlsxCellText(bytes, "Sheet1", "B1"));

        (uint StyleIndex, uint NumberFormatId, uint FontId, uint FillId, uint BorderId, string? FormatCode)
            dateStyle = OoxmlValidation.ReadXlsxCellStyle(bytes, "Sheet1", "A1");
        (uint StyleIndex, uint NumberFormatId, uint FontId, uint FillId, uint BorderId, string? FormatCode)
            numberStyle = OoxmlValidation.ReadXlsxCellStyle(bytes, "Sheet1", "B1");

        Assert.Equal("yyyy-mm-dd", dateStyle.FormatCode);
        Assert.Equal("#,##0.00", numberStyle.FormatCode);
        Assert.Equal(0U, dateStyle.FontId);
        Assert.Equal(0U, dateStyle.FillId);
        Assert.Equal(0U, dateStyle.BorderId);
    }

    [Fact]
    public void ProducedPackage_WithCellStyles_IsValidAndReadableByTheSdk()
    {
        CellStyle style = new()
        {
            FontName = "Arial",
            FontSizePoints = 12,
            Bold = true,
            FontColor = RgbColor.FromRgb(0x11, 0x22, 0x33),
            FillColor = RgbColor.FromRgb(0x44, 0x55, 0x66),
            Borders = new CellBorders(bottom: new CellBorder(CellBorderStyle.Thin)),
            HorizontalAlignment = CellHorizontalAlignment.Center,
            WrapText = true,
        };

        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet["A1"].Value = "Estilo";
            sheet["A1"].Style = style;
            sheet["B1"].Style = style;
            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));

        (uint StyleIndex, S.CellFormat Format, S.Font Font, S.Fill Fill, S.Border Border) first =
            OoxmlValidation.ReadXlsxCellStyleParts(bytes, "Sheet1", "A1");
        (uint StyleIndex, S.CellFormat Format, S.Font Font, S.Fill Fill, S.Border Border) second =
            OoxmlValidation.ReadXlsxCellStyleParts(bytes, "Sheet1", "B1");

        Assert.Equal(first.StyleIndex, second.StyleIndex);
        Assert.Equal("Arial", first.Font.FontName?.Val?.Value);
        Assert.Equal("FF112233", first.Font.Color?.Rgb?.Value);
        Assert.Equal("FF445566", first.Fill.PatternFill?.ForegroundColor?.Rgb?.Value);
        Assert.Equal(S.BorderStyleValues.Thin, first.Border.BottomBorder!.Style!.Value);
        Assert.Equal(S.HorizontalAlignmentValues.Center, first.Format.Alignment!.Horizontal!.Value);
        Assert.Equal("Estilo", OoxmlValidation.ReadXlsxCellText(bytes, "Sheet1", "A1"));
    }

    [Fact]
    public void OpenedWorkbook_WithSharedStringsAndMergedCells_ResolvesTextAndReportsDiagnostic()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithSharedStringsAndMergedCells("Hola", "Mundo");

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            Worksheet sheet = workbook.Worksheets[0];
            Assert.Equal("Hola", sheet["A1"].Value);
            Assert.Equal("Mundo", sheet["B1"].Value);

            OfficeEngineDiagnostic diagnostic = Assert.Single(workbook.Diagnostics);
            Assert.Equal(DiagnosticSeverity.Info, diagnostic.Severity);
            Assert.Contains("mergeCells", diagnostic.Message, StringComparison.Ordinal);

            produced = SaveBytes.ToBytes(workbook);
        }

        // The merged range that this version does not model must be preserved.
        string sheetXml = TestPackage.ReadEntryText(produced, "xl/worksheets/sheet1.xml");
        Assert.Contains("mergeCells", sheetXml, StringComparison.Ordinal);
        Assert.Contains("A1:B1", sheetXml, StringComparison.Ordinal);

        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void OpenedWorkbook_WithFormula_ReturnsCachedValueAndPreservesFormula()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormula("B1*2", "10");

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            // Only the stored value is available; this version does not calculate formulas.
            Assert.Equal(10.0, workbook.Worksheets[0]["A1"].Value);
            Assert.Equal("=B1*2", workbook.Worksheets[0]["A1"].Formula);

            produced = SaveBytes.ToBytes(workbook);
        }

        string sheetXml = TestPackage.ReadEntryText(produced, "xl/worksheets/sheet1.xml");
        Assert.Contains("B1*2", sheetXml, StringComparison.Ordinal);
        Assert.Contains(">10<", sheetXml, StringComparison.Ordinal);

        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void OpenedWorkbook_WithRowMissingIndex_ThrowsInvalidOfficeFileExceptionInsteadOfReturningNull()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithRowWithoutIndex("x");

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => _ = workbook.Worksheets[0]["A1"].Value);
    }

    [Fact]
    public void ProducedPackage_ListsWorksheetNamesInWorkbookXml()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Worksheets[0].Name = "Datos";
        _ = workbook.Worksheets.Add("Resumen");

        string workbookXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/workbook.xml");

        Assert.Contains("Datos", workbookXml, StringComparison.Ordinal);
        Assert.Contains("Resumen", workbookXml, StringComparison.Ordinal);
    }
}
