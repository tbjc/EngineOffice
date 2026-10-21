using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>
/// Index-based cell access: 1-based row/column indices must address exactly the same cell as the
/// corresponding A1 reference, and must reject indices outside the worksheet limits.
/// </summary>
public sealed class ExcelCellIndexTests
{
    [Fact]
    public void CellIndexer_UsesOneBasedRowAndColumn()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Cell cell = sheet[1, 1];

        Assert.Equal("A1", cell.Reference);
        Assert.Equal(1, cell.Row);
        Assert.Equal(1, cell.Column);
    }

    [Fact]
    public void CellHandle_ReportsTheRowAndColumnOfItsReference()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();

        Cell cell = workbook.Worksheets[0]["C2"];

        Assert.Equal(2, cell.Row);
        Assert.Equal(3, cell.Column);
    }

    [Fact]
    public void CellIndexer_AddressesTheSameCellAsTheA1Reference()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        sheet["B7"].Value = "siete";
        Assert.Equal("siete", sheet[7, 2].Value);

        sheet[2, 3].Value = 42;
        Assert.Equal(42.0, Assert.IsType<double>(sheet["C2"].Value));
        Assert.Equal("C2", sheet[2, 3].Reference);
    }

    [Fact]
    public void CellIndexer_RoundTripsThroughSaveAndOpen()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0][10, 5].Value = "E10";
            bytes = SaveBytes.ToBytes(workbook);
        }

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));

        Assert.Equal("E10", reopened.Worksheets[0][10, 5].Value);
        Assert.Equal("E10", reopened.Worksheets[0]["E10"].Value);
    }

    [Fact]
    public void CellIndexer_LastRowAndColumn_AddressesTheLastCell()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();

        Cell cell = workbook.Worksheets[0][1_048_576, 16_384];

        Assert.Equal("XFD1048576", cell.Reference);
        Assert.Equal(1_048_576, cell.Row);
        Assert.Equal(16_384, cell.Column);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(-1, 1)]
    [InlineData(1_048_577, 1)]
    [InlineData(1, 0)]
    [InlineData(1, -1)]
    [InlineData(1, 16_385)]
    public void CellIndexer_OutOfRange_ThrowsArgumentOutOfRange(int row, int column)
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentOutOfRangeException>(() => _ = sheet[row, column]);
    }

    [Fact]
    public void CellIndexer_OutOfRange_NamesTheRejectedIndex()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        ArgumentOutOfRangeException rowError = Assert.Throws<ArgumentOutOfRangeException>(() => _ = sheet[0, 1]);
        ArgumentOutOfRangeException columnError = Assert.Throws<ArgumentOutOfRangeException>(() => _ = sheet[1, 0]);

        Assert.Equal("row", rowError.ParamName);
        Assert.Equal("column", columnError.ParamName);
    }

    [Fact]
    public void CellIndexer_ReadingAMissingCell_DoesNotCreateIt()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Null(sheet[99, 26].Value);

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("Z99", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void CellIndexer_WritesCellsInColumnOrder()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        sheet[1, 3].Value = "C1";
        sheet[1, 1].Value = "A1";
        sheet[1, 2].Value = "B1";

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.True(
            sheetXml.IndexOf("A1", StringComparison.Ordinal) < sheetXml.IndexOf("B1", StringComparison.Ordinal));
        Assert.True(
            sheetXml.IndexOf("B1", StringComparison.Ordinal) < sheetXml.IndexOf("C1", StringComparison.Ordinal));
    }

    [Fact]
    public void CellIndexer_OnDisposedWorkbook_ThrowsObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = sheet[1, 1]);
    }
}
