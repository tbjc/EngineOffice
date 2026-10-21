using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>
/// Data validation rules: list, whole number and decimal rules written by this version, and the
/// rules of other workbooks, which are read and preserved.
/// </summary>
public sealed class ExcelDataValidationTests
{
    [Fact]
    public void AddDataValidation_WithAList_WritesTheRule()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            ExcelDataValidation validation = ExcelDataValidation.List("A1:A10", ["uno", "dos", "tres"]);
            workbook.Worksheets[0].AddDataValidation(validation);

            Assert.Equal("A1:A10", validation.Range);
            Assert.Equal(ExcelDataValidationType.List, validation.Type);
            Assert.Equal(3, validation.AllowedValues.Count);
            Assert.Equal("\"uno,dos,tres\"", validation.Formula1);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        IReadOnlyList<(string? Type, string? Sqref, string? Formula1, string? Formula2, string? Operator)> rules =
            OoxmlValidation.ReadXlsxDataValidations(bytes, "Sheet1");
        Assert.Single(rules);
        Assert.Equal("list", rules[0].Type);
        Assert.Equal("A1:A10", rules[0].Sqref);
        Assert.Equal("\"uno,dos,tres\"", rules[0].Formula1);

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        ExcelDataValidation read = reopened.Worksheets[0].DataValidations[0];
        Assert.Equal(ExcelDataValidationType.List, read.Type);
        Assert.Equal(["uno", "dos", "tres"], read.AllowedValues);
        Assert.Equal("A1:A10", read.Range);
    }

    [Theory]
    [InlineData(1d, 100d, "between", "1", "100")]
    [InlineData(1d, null, "greaterThanOrEqual", "1", null)]
    [InlineData(null, 100d, "lessThanOrEqual", "100", null)]
    public void AddDataValidation_WithANumber_WritesTheOperator(
        double? minimum,
        double? maximum,
        string expectedOperator,
        string expectedFormula1,
        string? expectedFormula2)
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0].AddDataValidation(
                ExcelDataValidation.WholeNumber("B1:B10", minimum, maximum));
            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        IReadOnlyList<(string? Type, string? Sqref, string? Formula1, string? Formula2, string? Operator)> rules =
            OoxmlValidation.ReadXlsxDataValidations(bytes, "Sheet1");
        Assert.Equal(expectedOperator, rules[0].Operator);
        Assert.Equal(expectedFormula1, rules[0].Formula1);
        Assert.Equal(expectedFormula2, rules[0].Formula2);
    }

    [Fact]
    public void AddDataValidation_WithADecimalRule_WritesTheType()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0].AddDataValidation(
                ExcelDataValidation.DecimalNumber("C1:C10", 0.5, 2.5));
            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        IReadOnlyList<(string? Type, string? Sqref, string? Formula1, string? Formula2, string? Operator)> rules =
            OoxmlValidation.ReadXlsxDataValidations(bytes, "Sheet1");
        Assert.Equal("decimal", rules[0].Type);
        Assert.Equal("0.5", rules[0].Formula1);
        Assert.Equal("2.5", rules[0].Formula2);
    }

    [Fact]
    public void AddDataValidation_WithAnInvalidRange_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => ExcelDataValidation.List("B2:A1", ["a"]));
        Assert.Throws<ArgumentException>(() => ExcelDataValidation.List("A", ["a"]));
        Assert.Throws<ArgumentException>(() => ExcelDataValidation.WholeNumber("Hoja1!A1"));
        Assert.Empty(sheet.DataValidations);
    }

    [Fact]
    public void AddDataValidation_WithAnInvalidList_Throws()
    {
        Assert.Throws<ArgumentException>(() => ExcelDataValidation.List("A1", []));
        Assert.Throws<ArgumentException>(() => ExcelDataValidation.List("A1", ["uno,dos"]));
        Assert.Throws<ArgumentException>(() => ExcelDataValidation.List("A1", ["con\"comilla"]));
        Assert.Throws<ArgumentException>(() => ExcelDataValidation.List("A1", [string.Empty]));
    }

    [Fact]
    public void AddDataValidation_WithTheMinimumGreaterThanTheMaximum_Throws()
    {
        Assert.Throws<ArgumentException>(() => ExcelDataValidation.WholeNumber("A1", 10, 1));
        Assert.Throws<ArgumentException>(() => ExcelDataValidation.DecimalNumber("A1", 10, 1));
    }

    [Fact]
    public void DataValidations_ReadTheRulesOfAnotherWorkbook()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithDataValidations(
            ("list", "A1:A10", "\"uno,dos\"", null),
            ("whole", "B1:B10", "1", "100"),
            ("custom", "C1:C10", "A1>0", null));

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        Assert.Equal(3, workbook.Worksheets[0].DataValidations.Count);
        Assert.Equal(ExcelDataValidationType.List, workbook.Worksheets[0].DataValidations[0].Type);
        Assert.Equal(["uno", "dos"], workbook.Worksheets[0].DataValidations[0].AllowedValues);
        Assert.Equal(ExcelDataValidationType.WholeNumber, workbook.Worksheets[0].DataValidations[1].Type);
        Assert.Equal("1", workbook.Worksheets[0].DataValidations[1].Formula1);
        Assert.Equal(ExcelDataValidationType.Custom, workbook.Worksheets[0].DataValidations[2].Type);
        Assert.Equal("A1>0", workbook.Worksheets[0].DataValidations[2].Formula1);
    }

    [Fact]
    public void DataValidations_ArePreservedWhenTheWorkbookIsSaved()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithDataValidations(
            ("list", "A1:A10", "\"uno,dos\"", null),
            ("custom", "C1:C10", "A1>0", null));

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            produced = SaveBytes.ToBytes(workbook);
        }

        IReadOnlyList<(string? Type, string? Sqref, string? Formula1, string? Formula2, string? Operator)> rules =
            OoxmlValidation.ReadXlsxDataValidations(produced, "Hoja1");
        Assert.Equal(2, rules.Count);
        Assert.Equal("list", rules[0].Type);
        Assert.Equal("custom", rules[1].Type);
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void RemoveDataValidation_RemovesTheRule()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.AddDataValidation(ExcelDataValidation.List("A1:A10", ["uno"]));
        sheet.AddDataValidation(ExcelDataValidation.WholeNumber("B1:B10", 1));

        Assert.True(sheet.RemoveDataValidation("A1:A10"));
        Assert.False(sheet.RemoveDataValidation("A1:A10"));
        Assert.Single(sheet.DataValidations);

        Assert.True(sheet.RemoveDataValidation("B1:B10"));
        Assert.Empty(sheet.DataValidations);

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("dataValidations", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void OpenedWorkbook_WithDataValidations_DoesNotReportThemAsUnmodelledContent()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithDataValidations(("list", "A1:A10", "\"uno\"", null));

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        Assert.Empty(workbook.Diagnostics);
    }

    [Fact]
    public void DisposedWorkbook_DataValidationsThrowsObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].DataValidations);
    }
}
