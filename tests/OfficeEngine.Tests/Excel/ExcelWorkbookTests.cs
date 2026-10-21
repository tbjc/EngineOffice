using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

public sealed class ExcelWorkbookTests
{
    [Fact]
    public void Create_HasSingleDefaultWorksheet()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();

        Assert.Single(workbook.Worksheets);
        Assert.Equal("Sheet1", workbook.Worksheets[0].Name);
        Assert.Empty(workbook.Diagnostics);
    }

    [Fact]
    public void Create_ThenEdit_ThenSave_Reopen_RoundTripsValuesAndTypes()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Datos";
            sheet["A1"].Value = "Texto";
            sheet["B2"].Value = 125.5;
            sheet["C3"].Value = true;
            sheet["D4"].Value = 0;
            bytes = SaveBytes.ToBytes(workbook);
        }

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Worksheet datos = reopened.Worksheets["datos"];

        Assert.Equal("Datos", datos.Name);
        Assert.Equal("Texto", Assert.IsType<string>(datos["A1"].Value));
        Assert.Equal(125.5, Assert.IsType<double>(datos["B2"].Value));
        Assert.True(Assert.IsType<bool>(datos["C3"].Value));
        Assert.Equal(0.0, Assert.IsType<double>(datos["D4"].Value));
        Assert.Null(datos["E5"].Value);
    }

    [Fact]
    public void Cell_ReadMissingCell_ReturnsNullWithoutCreatingIt()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Null(sheet["Z99"].Value);

        byte[] bytes = SaveBytes.ToBytes(workbook);
        string sheetXml = TestPackage.ReadEntryText(bytes, "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("Z99", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void Cell_AssignNull_ClearsExistingValue()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet["A1"].Value = "value";

        sheet["A1"].Value = null;

        Assert.Null(sheet["A1"].Value);
        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("value", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void Cell_AssignNull_DoesNotCreateAMissingCell()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        sheet["A1"].Value = null;

        Assert.Null(sheet["A1"].Value);
        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void Cell_WritesNumbersWithInvariantDecimalSeparator()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Worksheets[0]["A1"].Value = 1.25;

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.Contains("1.25", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void Cell_RejectsUnsupportedValueType()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<UnsupportedOfficeFeatureException>(() => sheet["A1"].Value = Guid.Empty);
    }

    [Theory]
    [InlineData("")]
    [InlineData("1A")]
    [InlineData("A0")]
    [InlineData("0")]
    [InlineData("A")]
    [InlineData("XFE1")]
    [InlineData("A1048577")]
    public void Cell_InvalidReference_ThrowsArgumentException(string reference)
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => _ = sheet[reference]);
    }

    [Fact]
    public void Worksheets_Add_AppendsAndIsAddressableByName()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        _ = workbook.Worksheets.Add("Resumen");

        Assert.Equal(2, workbook.Worksheets.Count);
        Assert.Equal("Resumen", workbook.Worksheets[1].Name);
        Assert.Equal("Sheet1", workbook.Worksheets["sheet1"].Name);
    }

    [Fact]
    public void Worksheets_Add_RejectsDuplicateAndInvalidNames()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();

        Assert.Throws<ArgumentException>(() => workbook.Worksheets.Add("sheet1"));
        Assert.Throws<ArgumentException>(() => workbook.Worksheets.Add("SHEET1"));
        Assert.Throws<ArgumentException>(() => workbook.Worksheets.Add(new string('a', 32)));
        Assert.Throws<ArgumentException>(() => workbook.Worksheets.Add("bad:name"));
        Assert.Throws<ArgumentNullException>(() => workbook.Worksheets.Add(null!));
    }

    [Fact]
    public void Worksheets_Rename_ValidatesUniquenessButAllowsCaseChange()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Worksheets.Add("Resumen");

        Assert.Throws<ArgumentException>(() => workbook.Worksheets["Resumen"].Name = "sheet1");

        workbook.Worksheets[0].Name = "SHEET1";
        Assert.Equal("SHEET1", workbook.Worksheets[0].Name);
    }

    [Fact]
    public void Worksheets_Remove_DeletesSheetButNotTheLastOne()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Worksheets.Add("Resumen");

        Assert.True(workbook.Worksheets.Remove("Resumen"));
        Assert.Single(workbook.Worksheets);
        Assert.False(workbook.Worksheets.Remove("Resumen"));
        Assert.Throws<InvalidOperationException>(() => workbook.Worksheets.Remove("Sheet1"));
    }

    [Fact]
    public void Worksheets_IndexByName_UnknownName_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Assert.Throws<ArgumentException>(() => _ = workbook.Worksheets["Nope"]);
    }

    [Fact]
    public void Save_ToFile_ThenOpen_RoundTrips()
    {
        string directory = TestData.NewTemporaryDirectory();
        string path = Path.Combine(directory, "book.xlsx");
        try
        {
            using (ExcelWorkbook workbook = ExcelWorkbook.Create())
            {
                workbook.Worksheets[0]["A1"].Value = "on disk";
                workbook.Save(path);
            }

            Assert.True(File.Exists(path));

            using (ExcelWorkbook reopened = ExcelWorkbook.Open(path))
            {
                Assert.Equal("on disk", reopened.Worksheets[0]["A1"].Value);
            }

            using (ExcelWorkbook again = ExcelWorkbook.Open(path))
            {
                again.Worksheets[0]["B1"].Value = 42;
                again.Save(path);
            }

            using (ExcelWorkbook third = ExcelWorkbook.Open(path))
            {
                Assert.Equal("on disk", third.Worksheets[0]["A1"].Value);
                Assert.Equal(42.0, third.Worksheets[0]["B1"].Value);
            }
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void DisposedWorkbook_ThrowsObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet["A1"].Value = 1;
        workbook.Dispose();

        using MemoryStream target = new();
        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets);
        Assert.Throws<ObjectDisposedException>(() => workbook.Save(target));
        Assert.Throws<ObjectDisposedException>(() => _ = sheet["A1"].Value);
    }

    [Fact]
    public void Open_RejectsNullStream() => Assert.Throws<ArgumentNullException>(() => ExcelWorkbook.Open((Stream)null!));

    [Fact]
    public void Save_RejectsInvalidArguments()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Assert.Throws<ArgumentNullException>(() => workbook.Save((Stream)null!));
        Assert.Throws<ArgumentException>(() => workbook.Save(string.Empty));

        using MemoryStream readOnly = new([], writable: false);
        Assert.Throws<ArgumentException>(() => workbook.Save(readOnly));
    }
}
