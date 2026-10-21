using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>Structured tables: the table part, its columns and the reference from the worksheet.</summary>
public sealed class ExcelTableTests
{
    [Fact]
    public void AddTable_CreatesTheTablePartAndTheReference()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet["A1"].Value = "Concepto";
            sheet["B1"].Value = "Importe";

            ExcelTable table = sheet.AddTable("Tabla1", "A1:B3");

            Assert.Equal("Tabla1", table.Name);
            Assert.Equal("A1:B3", table.Range);
            Assert.True(table.HasHeaderRow);
            Assert.Equal(["Concepto", "Importe"], table.ColumnNames);
            Assert.Single(sheet.Tables);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.True(TestPackage.HasEntry(bytes, "xl/tables/table1.xml"));

        IReadOnlyList<(string Name, string? Reference, uint HeaderRowCount, string[] Columns)> tables =
            OoxmlValidation.ReadXlsxTables(bytes, "Sheet1");
        Assert.Single(tables);
        Assert.Equal("Tabla1", tables[0].Name);
        Assert.Equal("A1:B3", tables[0].Reference);
        Assert.Equal(1U, tables[0].HeaderRowCount);
        Assert.Equal(["Concepto", "Importe"], tables[0].Columns);

        string sheetXml = TestPackage.ReadEntryText(bytes, "xl/worksheets/sheet1.xml");
        Assert.Contains("tableParts", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void AddTable_WithoutHeaderRow_GeneratesColumnNames()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            ExcelTable table = workbook.Worksheets[0].AddTable("Tabla1", "A1:B3", hasHeaderRow: false);

            Assert.False(table.HasHeaderRow);
            Assert.Equal(["Column1", "Column2"], table.ColumnNames);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.Equal(0U, OoxmlValidation.ReadXlsxTables(bytes, "Sheet1")[0].HeaderRowCount);
    }

    [Fact]
    public void AddTable_WithEmptyHeaderCells_GeneratesTheMissingColumnNames()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet["A1"].Value = "Concepto";

        ExcelTable table = sheet.AddTable("Tabla1", "A1:B3");

        Assert.Equal(["Concepto", "Column2"], table.ColumnNames);
    }

    [Fact]
    public void AddTable_WithAnInvalidName_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => sheet.AddTable("A1", "A1:B3"));
        Assert.Throws<ArgumentException>(() => sheet.AddTable("con espacio", "A1:B3"));
        Assert.Throws<ArgumentException>(() => sheet.AddTable("1abc", "A1:B3"));
        Assert.Throws<ArgumentException>(() => sheet.AddTable(string.Empty, "A1:B3"));
        Assert.Empty(sheet.Tables);
    }

    [Fact]
    public void AddTable_WithAnExistingName_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.AddTable("Tabla1", "A1:B3");

        Assert.Throws<ArgumentException>(() => sheet.AddTable("Tabla1", "A1:B3"));
        Assert.Single(sheet.Tables);
    }

    [Fact]
    public void AddTable_WithAnInvalidRange_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => sheet.AddTable("Tabla1", "B2:A1"));
        Assert.Empty(sheet.Tables);
    }

    [Fact]
    public void Tables_ReadsTheTablesOfAnotherWorkbook()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithTable("Datos", "A1:C5", ["Concepto", "Importe", "Total"]);

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        ExcelTable table = workbook.Worksheets[0].Tables[0];
        Assert.Equal("Datos", table.Name);
        Assert.Equal("A1:C5", table.Range);
        Assert.True(table.HasHeaderRow);
        Assert.Equal(["Concepto", "Importe", "Total"], table.ColumnNames);
    }

    [Fact]
    public void Tables_ArePreservedWhenTheWorkbookIsSaved()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithTable("Datos", "A1:C5", ["Concepto", "Importe"]);

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.True(TestPackage.HasEntry(produced, "xl/tables/table1.xml"));
        Assert.Single(OoxmlValidation.ReadXlsxTables(produced, "Hoja1"));
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void RemoveTable_RemovesTheTablePart()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.AddTable("Tabla1", "A1:B3");

            Assert.True(sheet.RemoveTable("Tabla1"));
            Assert.False(sheet.RemoveTable("Tabla1"));
            Assert.Empty(sheet.Tables);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.False(TestPackage.HasEntry(bytes, "xl/tables/table1.xml"));
        string sheetXml = TestPackage.ReadEntryText(bytes, "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("tableParts", sheetXml, StringComparison.Ordinal);
        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void OpenedWorkbook_WithATable_DoesNotReportItAsUnmodelledContent()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithTable("Datos", "A1:C5", ["Concepto"]);

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        Assert.Empty(workbook.Diagnostics);
    }

    [Fact]
    public void DisposedWorkbook_TablesThrowsObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].Tables);
    }
}
