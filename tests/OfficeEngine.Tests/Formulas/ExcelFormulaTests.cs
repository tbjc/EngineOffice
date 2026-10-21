using OfficeEngine.Excel;
using OfficeEngine.Formulas;

namespace OfficeEngine.Tests.Formulas;

/// <summary>The engine reading the values of a workbook.</summary>
public sealed class ExcelFormulaTests
{
    [Fact]
    public void EvaluateFormula_ReadsTheValuesOfTheWorkbook()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Worksheets[0]["A1"].Value = 6d;
        workbook.Worksheets[0]["A2"].Value = 7d;

        object? result = workbook.EvaluateFormula("=A1*A2");

        Assert.Equal(42d, result);
    }

    [Fact]
    public void EvaluateFormula_UsesTheSheetItIsGiven()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Worksheets.Add("Datos");
        workbook.Worksheets["Datos"]["A1"].Value = 3d;

        object? result = workbook.EvaluateFormula("=A1*2", "Datos");

        Assert.Equal(6d, result);
    }

    [Fact]
    public void EvaluateFormula_ReportsTheErrorOfExcel()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();

        object? result = workbook.EvaluateFormula("=1/0");

        FormulaError error = Assert.IsType<FormulaError>(result);
        Assert.Equal("#DIV/0!", error.Code);
    }

    [Fact]
    public void EvaluateFormula_WithAnEmptyCell_TreatsItAsZero()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();

        Assert.Equal(1d, workbook.EvaluateFormula("=A1+1"));
        Assert.Equal("x", workbook.EvaluateFormula("=A1&\"x\""));
    }

    [Fact]
    public void EvaluateFormula_ReadsTheTextOfACell()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Worksheets[0]["A1"].Value = "hola";

        Assert.Equal("hola!", workbook.EvaluateFormula("=A1&\"!\""));
    }

    [Fact]
    public void EvaluateFormula_WithTheSpanishNameOfAFunction_ComputesTheSameValue()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Worksheets[0]["A1"].Value = 6d;
        workbook.Worksheets[0]["A2"].Value = 7d;

        Assert.Equal(13d, workbook.EvaluateFormula("=SUMA(A1,A2)"));
    }

    [Fact]
    public void Formula_WithTheSpanishNameOfAFunction_IsStoredWithTheCanonicalName()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();

        workbook.Worksheets[0]["B1"].Formula = "=SUMA(A1:A2)";

        // Excel stores the names of the functions in English.
        Assert.Equal("=SUM(A1:A2)", workbook.Worksheets[0]["B1"].Formula);
    }
}
