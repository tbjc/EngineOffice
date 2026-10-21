using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>
/// Rectangular ranges: both corners, row-major iteration and bulk values. Reading a missing cell
/// must yield null without creating it; writing must not create cells for null entries.
/// </summary>
public sealed class ExcelCellRangeTests
{
    [Fact]
    public void GetRange_ParsesBothCornersAndReportsTheReference()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        CellRange range = workbook.Worksheets[0].GetRange("B2:D5");

        Assert.Equal("B2:D5", range.Reference);
        Assert.Equal(2, range.FirstRow);
        Assert.Equal(2, range.FirstColumn);
        Assert.Equal(5, range.LastRow);
        Assert.Equal(4, range.LastColumn);
        Assert.Equal(4, range.RowCount);
        Assert.Equal(3, range.ColumnCount);
        Assert.Equal("B2:D5", range.ToString());
    }

    [Fact]
    public void GetRange_NormalizesLowercaseReferences()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();

        CellRange range = workbook.Worksheets[0].GetRange("b2:d5");

        Assert.Equal("B2:D5", range.Reference);
    }

    [Fact]
    public void GetRange_IntOverload_BuildsTheSameRange()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();

        CellRange range = workbook.Worksheets[0].GetRange(2, 2, 5, 4);

        Assert.Equal("B2:D5", range.Reference);
        Assert.Equal(2, range.FirstRow);
        Assert.Equal(2, range.FirstColumn);
        Assert.Equal(5, range.LastRow);
        Assert.Equal(4, range.LastColumn);
    }

    [Fact]
    public void Cells_EnumeratesInRowMajorOrder()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();

        string[] references = workbook.Worksheets[0].GetRange("B2:C3").Cells
            .Select(cell => cell.Reference)
            .ToArray();

        Assert.Collection(
            references,
            reference => Assert.Equal("B2", reference),
            reference => Assert.Equal("C2", reference),
            reference => Assert.Equal("B3", reference),
            reference => Assert.Equal("C3", reference));
    }

    [Fact]
    public void Cells_IsALiveView()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        CellRange range = sheet.GetRange("A1:B1");

        sheet["A1"].Value = "antes";
        Assert.Equal("antes", range.Cells.First().Value);

        sheet["A1"].Value = "después";
        Assert.Equal("después", range.Cells.First().Value);
    }

    [Fact]
    public void Cells_ReadingDoesNotCreateMissingCells()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        _ = sheet.GetRange("Z99:AA100").Cells.ToList();

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("Z99", sheetXml, StringComparison.Ordinal);
        Assert.DoesNotContain("AA100", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void GetValues_ReadsCellsInRowMajorOrderAndReturnsNullForMissingOnes()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet["B2"].Value = "b2";
        sheet["C3"].Value = 3.0;

        object?[,] values = sheet.GetRange("B2:C3").GetValues();

        Assert.Equal(2, values.GetLength(0));
        Assert.Equal(2, values.GetLength(1));
        Assert.Equal("b2", values[0, 0]);
        Assert.Null(values[0, 1]);
        Assert.Null(values[1, 0]);
        Assert.Equal(3.0, Assert.IsType<double>(values[1, 1]));
    }

    [Fact]
    public void SetValues_WritesEveryCellAndGetValuesReadsItBack()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        CellRange range = sheet.GetRange("A1:B2");

        range.SetValues(new object?[,]
        {
            { "texto", 1.5 },
            { true, null },
        });

        Assert.Equal("texto", sheet["A1"].Value);
        Assert.Equal(1.5, Assert.IsType<double>(sheet["B1"].Value));
        Assert.True(Assert.IsType<bool>(sheet["A2"].Value));
        Assert.Null(sheet["B2"].Value);

        object?[,] values = range.GetValues();
        Assert.Equal("texto", values[0, 0]);
        Assert.Equal(1.5, Assert.IsType<double>(values[0, 1]));
        Assert.True(Assert.IsType<bool>(values[1, 0]));
        Assert.Null(values[1, 1]);
    }

    [Fact]
    public void SetValues_ClearsExistingCellsWithoutCreatingMissingOnes()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet["A1"].Value = "uno";
        sheet["B1"].Value = "dos";

        sheet.GetRange("A1:B2").SetValues(new object?[2, 2]);

        Assert.Null(sheet["A1"].Value);
        Assert.Null(sheet["B1"].Value);
        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A2", sheetXml, StringComparison.Ordinal);
        Assert.DoesNotContain("B2", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void SetValues_RequiresExactDimensions()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        CellRange range = workbook.Worksheets[0].GetRange("A1:B2");

        Assert.Throws<ArgumentException>(() => range.SetValues(new object?[1, 2]));
        Assert.Throws<ArgumentException>(() => range.SetValues(new object?[2, 1]));
        Assert.Throws<ArgumentNullException>(() => range.SetValues(null!));
    }

    [Fact]
    public void SetValues_UnsupportedValue_LeavesTheRangeUnchanged()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        CellRange range = sheet.GetRange("A1:B1");

        Assert.Throws<UnsupportedOfficeFeatureException>(() => range.SetValues(new object?[,]
        {
            { "ok", Guid.Empty },
        }));

        Assert.Null(sheet["A1"].Value);
        Assert.Null(sheet["B1"].Value);
    }

    [Fact]
    public void RangeValues_RoundTripThroughSaveAndOpen()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0].GetRange("B2:C3").SetValues(new object?[,]
            {
                { "a", 2.5 },
                { false, null },
            });
            bytes = SaveBytes.ToBytes(workbook);
        }

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        object?[,] values = reopened.Worksheets[0].GetRange("B2:C3").GetValues();

        Assert.Equal("a", values[0, 0]);
        Assert.Equal(2.5, Assert.IsType<double>(values[0, 1]));
        Assert.False(Assert.IsType<bool>(values[1, 0]));
        Assert.Null(values[1, 1]);
    }

    [Theory]
    [InlineData("")]
    [InlineData("A1")]
    [InlineData("A1:")]
    [InlineData(":B2")]
    [InlineData("A1:B2:C3")]
    [InlineData("A0:B2")]
    [InlineData("B2:A1")]
    [InlineData("A1:B2 ")]
    [InlineData("$A$1:$B$2")]
    public void GetRange_InvalidReference_ThrowsArgumentException(string reference)
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => _ = sheet.GetRange(reference));
    }

    [Fact]
    public void GetRange_NullReference_ThrowsArgumentNullException()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();

        Assert.Throws<ArgumentNullException>(() => _ = workbook.Worksheets[0].GetRange(null!));
    }

    [Theory]
    [InlineData(0, 1, 1, 1)]
    [InlineData(1, 0, 1, 1)]
    [InlineData(1, 1, 1_048_577, 1)]
    [InlineData(1, 1, 1, 16_385)]
    public void GetRange_IntOverload_OutOfRange_ThrowsArgumentOutOfRange(
        int firstRow,
        int firstColumn,
        int lastRow,
        int lastColumn)
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentOutOfRangeException>(
            () => sheet.GetRange(firstRow, firstColumn, lastRow, lastColumn));
    }

    [Fact]
    public void GetRange_IntOverload_NamesTheRejectedCorner()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        ArgumentOutOfRangeException firstRowError =
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.GetRange(0, 1, 1, 1));
        ArgumentOutOfRangeException lastColumnError =
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.GetRange(1, 1, 1, 16_385));

        Assert.Equal("firstRow", firstRowError.ParamName);
        Assert.Equal("lastColumn", lastColumnError.ParamName);
    }

    [Fact]
    public void GetRange_IntOverload_ReversedCorners_ThrowsArgumentException()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => sheet.GetRange(2, 1, 1, 1));
        Assert.Throws<ArgumentException>(() => sheet.GetRange(1, 2, 1, 1));
    }

    [Fact]
    public void GetValues_ReturnsDateTimeForDateFormattedCells()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        DateTime expected = new(2024, 1, 1);
        sheet["A1"].SetDateTime(expected, "yyyy-mm-dd");

        object?[,] values = sheet.GetRange("A1:B1").GetValues();

        Assert.Equal(expected, Assert.IsType<DateTime>(values[0, 0]));
        Assert.Null(values[0, 1]);
    }

    [Fact]
    public void SetValues_DateTime_LeavesTheRangeUnchanged()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        CellRange range = sheet.GetRange("A1:B1");

        Assert.Throws<UnsupportedOfficeFeatureException>(() => range.SetValues(new object?[,]
        {
            { "ok", new DateTime(2024, 1, 1) },
        }));

        Assert.Null(sheet["A1"].Value);
        Assert.Null(sheet["B1"].Value);
    }

    [Fact]
    public void Range_OnDisposedWorkbook_ThrowsObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        CellRange range = sheet.GetRange("A1:B2");
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => sheet.GetRange("A1:B2"));
        Assert.Throws<ObjectDisposedException>(() => _ = range.Cells);
        Assert.Throws<ObjectDisposedException>(() => range.GetValues());
        Assert.Throws<ObjectDisposedException>(() => range.SetValues(new object?[2, 2]));
    }
}
