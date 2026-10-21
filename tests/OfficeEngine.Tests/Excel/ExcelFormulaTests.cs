using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>
/// Formula writing is separate from calculation: a formula is stored as a formula, any cached result
/// is removed when the formula changes, and the library never evaluates it.
/// </summary>
public sealed class ExcelFormulaTests
{
    [Fact]
    public void Formula_WriteWithoutCachedValue_RoundTrips()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet["A1"].Value = 10;
            sheet["B1"].Value = 2;
            sheet["C1"].Formula = "=A1*B1";

            Assert.Equal("=A1*B1", sheet["C1"].Formula);
            Assert.Null(sheet["C1"].Value);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.Equal("A1*B1", OoxmlValidation.ReadXlsxCellFormula(bytes, "Sheet1", "C1").Formula);
        Assert.Null(OoxmlValidation.ReadXlsxCellFormula(bytes, "Sheet1", "C1").CachedValue);
        Assert.True(OoxmlValidation.ReadXlsxFullCalculationOnLoad(bytes));

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Cell cell = reopened.Worksheets[0]["C1"];
        Assert.Equal("=A1*B1", cell.Formula);
        Assert.Null(cell.Value);
    }

    [Fact]
    public void Formula_WithoutLeadingEquals_IsNormalized()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            cell.Formula = "B1*2";

            Assert.Equal("=B1*2", cell.Formula);
            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Equal("B1*2", OoxmlValidation.ReadXlsxCellFormula(bytes, "Sheet1", "A1").Formula);
        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void Formula_DoesNotEvaluate()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            cell.Formula = "=1+1";

            Assert.Null(cell.Value);
            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Null(OoxmlValidation.ReadXlsxCellFormula(bytes, "Sheet1", "A1").CachedValue);
        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void Formula_Null_RemovesFormulaAndCachedValue()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormula("B1*2", "10");

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            Assert.Equal("=B1*2", cell.Formula);
            Assert.Equal(10.0, cell.Value);

            cell.Formula = null;

            Assert.Null(cell.Formula);
            Assert.Null(cell.Value);
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Null(OoxmlValidation.ReadXlsxCellFormula(produced, "Hoja1", "A1").Formula);
        Assert.Null(OoxmlValidation.ReadXlsxCellFormula(produced, "Hoja1", "A1").CachedValue);
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void Formula_Set_ClearsStaleCachedValue()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormula("B1*2", "10");

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            cell.Formula = "=B1*3";

            Assert.Equal("=B1*3", cell.Formula);
            Assert.Null(cell.Value);
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Equal("B1*3", OoxmlValidation.ReadXlsxCellFormula(produced, "Hoja1", "A1").Formula);
        Assert.Null(OoxmlValidation.ReadXlsxCellFormula(produced, "Hoja1", "A1").CachedValue);
        Assert.True(OoxmlValidation.ReadXlsxFullCalculationOnLoad(produced));
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void Value_SetOnFormulaCell_RemovesFormula()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormula("B1*2", "10");

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            cell.Value = 42;

            Assert.Null(cell.Formula);
            Assert.Equal(42.0, cell.Value);
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Null(OoxmlValidation.ReadXlsxCellFormula(produced, "Hoja1", "A1").Formula);
        Assert.Equal("42", OoxmlValidation.ReadXlsxCellFormula(produced, "Hoja1", "A1").CachedValue);
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void Value_NullOnFormulaCell_RemovesFormulaAndCachedValue()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormula("B1*2", "10");

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            cell.Value = null;

            Assert.Null(cell.Formula);
            Assert.Null(cell.Value);
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Null(OoxmlValidation.ReadXlsxCellFormula(produced, "Hoja1", "A1").Formula);
        Assert.Null(OoxmlValidation.ReadXlsxCellFormula(produced, "Hoja1", "A1").CachedValue);
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void Formula_RequestsRecalculationAfterExistingDefinedNames()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormula(
            "B1*2",
            "10",
            includeDefinedName: true,
            includeTrailingElement: true);

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            workbook.Worksheets[0]["A1"].Formula = "=B1*3";
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.True(OoxmlValidation.ReadXlsxFullCalculationOnLoad(produced));
        string workbookXml = TestPackage.ReadEntryText(produced, "xl/workbook.xml");
        int definedNames = workbookXml.IndexOf("definedNames", StringComparison.Ordinal);
        int calcPr = workbookXml.IndexOf("calcPr", StringComparison.Ordinal);
        int recovery = workbookXml.IndexOf("fileRecoveryPr", StringComparison.Ordinal);
        Assert.True(definedNames >= 0 && calcPr > definedNames && recovery > calcPr,
            "calcPr must come after definedNames and before fileRecoveryPr.");
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void Formula_Invalid_DoesNotCreateOrChangeTheCell()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Cell cell = workbook.Worksheets[0]["A1"];

        Assert.Throws<ArgumentException>(() => cell.Formula = string.Empty);
        Assert.Throws<ArgumentException>(() => cell.Formula = "=");
        Assert.Throws<ArgumentException>(() => cell.Formula = new string('A', 8193));
        Assert.Throws<ArgumentException>(() => cell.Formula = "=A1\u0001");
        cell.Formula = null;

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void Formula_InvalidOnExistingCell_LeavesFormulaAndCacheUnchanged()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormula("B1*2", "10");

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));
        Cell cell = workbook.Worksheets[0]["A1"];

        Assert.Throws<ArgumentException>(() => cell.Formula = "=");

        Assert.Equal("=B1*2", cell.Formula);
        Assert.Equal(10.0, cell.Value);
    }

    [Fact]
    public void Formula_PreservesStyleAndNumberFormat()
    {
        RgbColor fill = RgbColor.FromRgb(0x12, 0x34, 0x56);
        CellStyle style = new() { Bold = true, FillColor = fill };

        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            cell.SetDateTime(new DateTime(2024, 1, 31), "yyyy-mm-dd");
            cell.Style = style;
            cell.Formula = "=B1+1";

            Assert.Equal("yyyy-mm-dd", cell.NumberFormat);
            Assert.Equal(style, cell.Style);
            Assert.Null(cell.Value);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Cell reopenedCell = reopened.Worksheets[0]["A1"];
        Assert.Equal("=B1+1", reopenedCell.Formula);
        Assert.Equal("yyyy-mm-dd", reopenedCell.NumberFormat);
        Assert.Equal(style, reopenedCell.Style);
    }

    [Fact]
    public void DisposedWorkbook_FormulaThrowsObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        Cell cell = workbook.Worksheets[0]["A1"];
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = cell.Formula);
        Assert.Throws<ObjectDisposedException>(() => cell.Formula = "=A1");
    }
}
