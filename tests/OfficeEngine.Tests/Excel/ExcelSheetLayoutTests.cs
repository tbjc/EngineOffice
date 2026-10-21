using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>
/// The layout of a sheet: frozen panes, row height and hidden rows, column width and hidden
/// columns.
/// </summary>
public sealed class ExcelSheetLayoutTests
{
    [Fact]
    public void FreezePanes_Set_WritesThePane()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.FreezePanes = new ExcelFreezePanes(rows: 1, columns: 2);

            Assert.Equal(1, sheet.FreezePanes!.Rows);
            Assert.Equal(2, sheet.FreezePanes.Columns);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        (double? XSplit, double? YSplit, string? TopLeftCell, string? State, string? ActivePane) = OoxmlValidation.ReadXlsxPane(bytes, "Sheet1")!.Value;
        Assert.Equal(2d, XSplit);
        Assert.Equal(1d, YSplit);
        Assert.Equal("C2", TopLeftCell);
        Assert.Equal("frozen", State);

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        ExcelFreezePanes reopenedPanes = reopened.Worksheets[0].FreezePanes!;
        Assert.Equal(1, reopenedPanes.Rows);
        Assert.Equal(2, reopenedPanes.Columns);
    }

    [Theory]
    [InlineData(1, 0, "A2", "bottomLeft")]
    [InlineData(0, 2, "C1", "topRight")]
    public void FreezePanes_WithOnlyOneDirection_UsesTheRightCorner(
        int rows,
        int columns,
        string expectedTopLeftCell,
        string expectedActivePane)
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0].FreezePanes = new ExcelFreezePanes(rows, columns);
            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        (double? XSplit, double? YSplit, string? TopLeftCell, string? State, string? ActivePane) = OoxmlValidation.ReadXlsxPane(bytes, "Sheet1")!.Value;
        Assert.Equal(expectedTopLeftCell, TopLeftCell);
        Assert.Equal("frozen", State);
        Assert.Equal(expectedActivePane, ActivePane);
        if (columns == 0)
        {
            Assert.Null(XSplit);
        }
        else
        {
            Assert.Equal((double)columns, XSplit);
        }

        if (rows == 0)
        {
            Assert.Null(YSplit);
        }
        else
        {
            Assert.Equal((double)rows, YSplit);
        }
    }

    [Fact]
    public void FreezePanes_Null_RemovesIt()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.FreezePanes = new ExcelFreezePanes(1, 0);

        sheet.FreezePanes = null;

        Assert.Null(sheet.FreezePanes);
        Assert.Null(OoxmlValidation.ReadXlsxPane(SaveBytes.ToBytes(workbook), "Sheet1"));
    }

    [Fact]
    public void FreezePanes_ReadsThePaneOfAnotherWorkbook()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithSheetLayout(
            frozenColumns: 2,
            frozenRows: 3,
            topLeftCell: "C4");

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        ExcelFreezePanes panes = workbook.Worksheets[0].FreezePanes!;
        Assert.Equal(3, panes.Rows);
        Assert.Equal(2, panes.Columns);
    }

    [Fact]
    public void FreezePanes_WithInvalidValues_Throws()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelFreezePanes(-1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelFreezePanes(0, -1));
        Assert.Throws<ArgumentException>(() => new ExcelFreezePanes(0, 0));
    }

    [Fact]
    public void SetRowHeight_WritesTheHeight()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.SetRowHeight(2, 30.5);

            Assert.Equal(30.5, sheet.RowHeight(2));
            Assert.Null(sheet.RowHeight(3));

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.Equal(30.5, OoxmlValidation.ReadXlsxRow(bytes, 2U, "Sheet1").Height);
    }

    [Fact]
    public void SetRowHeight_WithAnInvalidHeight_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetRowHeight(1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetRowHeight(1, -5));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetRowHeight(1, 500));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetRowHeight(0, 10));
    }

    [Fact]
    public void SetRowHidden_HidesAndShowsTheRow()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.SetRowHidden(3, true);

            Assert.True(sheet.IsRowHidden(3));
            Assert.False(sheet.IsRowHidden(4));

            sheet.SetRowHidden(3, false);
            Assert.False(sheet.IsRowHidden(3));

            sheet.SetRowHidden(3, true);
            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.True(OoxmlValidation.ReadXlsxRow(bytes, 3U, "Sheet1").Hidden);
    }

    [Fact]
    public void SetColumnWidth_WritesTheWidth()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.SetColumnWidth(2, 12.5);

            Assert.Equal(12.5, sheet.ColumnWidth(2));
            Assert.Null(sheet.ColumnWidth(3));

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.Equal(12.5, OoxmlValidation.ReadXlsxColumn(bytes, 2, "Sheet1").Width);
    }

    [Fact]
    public void SetColumnWidth_WithAnInvalidWidth_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(1, -3));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(1, 300));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(0, 10));
    }

    [Fact]
    public void SetColumnHidden_HidesAndShowsTheColumn()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.SetColumnHidden(4, true);

            Assert.True(sheet.IsColumnHidden(4));
            Assert.False(sheet.IsColumnHidden(5));

            sheet.SetColumnHidden(4, false);
            Assert.False(sheet.IsColumnHidden(4));

            sheet.SetColumnHidden(4, true);
            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.True(OoxmlValidation.ReadXlsxColumn(bytes, 4, "Sheet1").Hidden);
    }

    [Fact]
    public void SetColumnWidth_OnAColumnGroupedWithOthers_SplitsTheElement()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithSheetLayout(
            columnIndex: 1,
            columnMax: 3,
            columnWidth: 10);

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            Worksheet sheet = workbook.Worksheets[0];
            Assert.Equal(10, sheet.ColumnWidth(2));

            sheet.SetColumnWidth(2, 20);

            Assert.Equal(10, sheet.ColumnWidth(1));
            Assert.Equal(20, sheet.ColumnWidth(2));
            Assert.Equal(10, sheet.ColumnWidth(3));

            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
        Assert.Equal(20, OoxmlValidation.ReadXlsxColumn(produced, 2, "Hoja1").Width);
        Assert.Equal(10, OoxmlValidation.ReadXlsxColumn(produced, 3, "Hoja1").Width);
    }

    [Fact]
    public void OpenedWorkbook_WithALayout_DoesNotReportItAsUnmodelledContent()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithSheetLayout(
            frozenColumns: 1,
            frozenRows: 1,
            topLeftCell: "B2",
            rowHeight: 20,
            columnWidth: 12);

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        Assert.Empty(workbook.Diagnostics);
    }

    [Fact]
    public void DisposedWorkbook_LayoutThrowsObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].FreezePanes);
        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].RowHeight(1));
        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].ColumnWidth(1));
    }
}
