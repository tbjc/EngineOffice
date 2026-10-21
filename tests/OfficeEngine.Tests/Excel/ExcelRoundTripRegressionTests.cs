using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>A round-trip over a generated workbook: values, styles and formulas must survive.</summary>
public sealed class ExcelRoundTripRegressionTests
{
    [Fact]
    public void AGeneratedWorkbook_RoundTripsWithoutLosingValues()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        for (int row = 1; row <= 500; row++)
        {
            sheet[row, 1].Value = $"Texto {row}";
            sheet[row, 2].Value = row * 1.5d;
            sheet[row, 3].Value = row % 2 == 0;
        }

        sheet["E1"].NumberFormat = "0.00";
        sheet["E2"].Formula = "=SUM(B1:B10)";

        using MemoryStream stream = new();
        workbook.Save(stream);
        stream.Position = 0;

        using ExcelWorkbook reopened = ExcelWorkbook.Open(stream);
        Worksheet reopenedSheet = reopened.Worksheets[0];
        Assert.Equal("Texto 500", reopenedSheet[500, 1].Value);
        Assert.Equal(750d, reopenedSheet[500, 2].Value);
        Assert.Equal(true, reopenedSheet[500, 3].Value);
        Assert.Equal("0.00", reopenedSheet["E1"].NumberFormat);
        Assert.Equal("=SUM(B1:B10)", reopenedSheet["E2"].Formula);
    }
}
