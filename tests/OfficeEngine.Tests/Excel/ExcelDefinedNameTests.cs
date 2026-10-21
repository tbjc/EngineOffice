using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>
/// Named ranges: the catalogue of defined names, the creation of workbook-scoped and sheet-scoped
/// names, and the preservation of the names this version does not interpret.
/// </summary>
public sealed class ExcelDefinedNameTests
{
    [Fact]
    public void AddNamedRange_CreatesWorkbookScopedName()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            ExcelDefinedName definedName = workbook.Worksheets[0].AddNamedRange("Datos", "A1:B2");

            Assert.Equal("Datos", definedName.Name);
            Assert.Equal("Sheet1!$A$1:$B$2", definedName.RefersTo);
            Assert.True(definedName.IsRange);
            Assert.Equal("A1:B2", definedName.Range);
            Assert.Null(definedName.Scope);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        IReadOnlyList<(string Name, string Text, uint? LocalSheetId)> names =
            OoxmlValidation.ReadXlsxDefinedNames(bytes);
        Assert.Single(names);
        Assert.Equal("Datos", names[0].Name);
        Assert.Equal("Sheet1!$A$1:$B$2", names[0].Text);
        Assert.Null(names[0].LocalSheetId);

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Assert.Single(reopened.DefinedNames);
        Assert.Equal("A1:B2", reopened.DefinedNames[0].Range);
        Assert.Null(reopened.DefinedNames[0].Scope);
    }

    [Fact]
    public void AddLocalNamedRange_CreatesSheetScopedName()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            ExcelDefinedName definedName = workbook.Worksheets[0].AddLocalNamedRange("Local", "A1:B2");

            Assert.Equal("Local", definedName.Name);
            Assert.Equal("$A$1:$B$2", definedName.RefersTo);
            Assert.True(definedName.IsRange);
            Assert.Equal("A1:B2", definedName.Range);
            Assert.NotNull(definedName.Scope);
            Assert.Equal("Sheet1", definedName.Scope!.Name);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        IReadOnlyList<(string Name, string Text, uint? LocalSheetId)> names =
            OoxmlValidation.ReadXlsxDefinedNames(bytes);
        Assert.Single(names);
        Assert.Equal("$A$1:$B$2", names[0].Text);
        Assert.Equal(0U, names[0].LocalSheetId);
    }

    [Fact]
    public void AddNamedRange_WithTheSameName_ReplacesTheTarget()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0].AddNamedRange("Datos", "A1:B2");
            ExcelDefinedName replaced = workbook.Worksheets[0].AddNamedRange("Datos", "C3");

            Assert.Equal("Sheet1!$C$3", replaced.RefersTo);
            Assert.Single(workbook.DefinedNames);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Single(OoxmlValidation.ReadXlsxDefinedNames(bytes));
    }

    [Fact]
    public void AddNamedRange_WithASingleCell_StoresThatCell()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            ExcelDefinedName definedName = workbook.Worksheets[0].AddNamedRange("Celda", "C3");

            Assert.Equal("Sheet1!$C$3", definedName.RefersTo);
            Assert.Equal("C3", definedName.Range);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void AddNamedRange_QuotesTheSheetNameWhenItNeedsIt()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets.Add("Mi Hoja");
            ExcelDefinedName definedName = sheet.AddNamedRange("Datos", "A1");

            Assert.Equal("'Mi Hoja'!$A$1", definedName.RefersTo);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.Equal("'Mi Hoja'!$A$1", OoxmlValidation.ReadXlsxDefinedNames(bytes)[0].Text);
    }

    [Fact]
    public void AddNamedRange_WithAnInvalidName_ThrowsAndChangesNothing()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => sheet.AddNamedRange(string.Empty, "A1"));
        Assert.Throws<ArgumentException>(() => sheet.AddNamedRange("A1", "B2"));
        Assert.Throws<ArgumentException>(() => sheet.AddNamedRange("1abc", "B2"));
        Assert.Throws<ArgumentException>(() => sheet.AddNamedRange("con*asterisco", "B2"));
        Assert.Throws<ArgumentException>(() => sheet.AddNamedRange(new string('a', 256), "B2"));

        Assert.Empty(workbook.DefinedNames);
        string workbookXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/workbook.xml");
        Assert.DoesNotContain("definedNames", workbookXml, StringComparison.Ordinal);
    }

    [Fact]
    public void AddNamedRange_WithAnInvalidRange_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => sheet.AddNamedRange("Datos", "B2:A1"));
        Assert.Throws<ArgumentException>(() => sheet.AddNamedRange("Datos", "A"));
        Assert.Throws<ArgumentException>(() => sheet.AddNamedRange("Datos", "Hoja1!A1"));
        Assert.Empty(workbook.DefinedNames);
    }

    [Fact]
    public void RemoveNamedRange_RemovesTheName()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Worksheets[0].AddNamedRange("Datos", "A1:B2");

        Assert.True(workbook.RemoveNamedRange("datos"));
        Assert.False(workbook.RemoveNamedRange("datos"));
        Assert.Empty(workbook.DefinedNames);

        string workbookXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/workbook.xml");
        Assert.DoesNotContain("definedNames", workbookXml, StringComparison.Ordinal);
    }

    [Fact]
    public void RemoveLocalNamedRange_RemovesTheNameOfThatWorksheetOnly()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet first = workbook.Worksheets[0];
        Worksheet second = workbook.Worksheets.Add("Otra");

        first.AddLocalNamedRange("Local", "A1");
        second.AddLocalNamedRange("Local", "B2");

        Assert.True(first.RemoveLocalNamedRange("Local"));
        Assert.Single(workbook.DefinedNames);
        Assert.Equal("Otra", workbook.DefinedNames[0].Scope!.Name);
    }

    [Fact]
    public void DefinedNames_ReadsTheNamesOfAnotherWorkbook()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithDefinedNames(
            ("Datos", "Hoja1!$A$1:$B$2", null),
            ("Local", "$C$1", 0U),
            ("Constante", "42", null),
            ("_xlnm.Print_Area", "Hoja1!$A$1:$A$5", 0U));

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        Assert.Equal(4, workbook.DefinedNames.Count);

        ExcelDefinedName datos = workbook.FindDefinedName("datos")!;
        Assert.NotNull(datos);
        Assert.True(datos.IsRange);
        Assert.Equal("A1:B2", datos.Range);
        Assert.Null(datos.Scope);

        ExcelDefinedName local = workbook.FindDefinedName("Local")!;
        Assert.NotNull(local);
        Assert.True(local.IsRange);
        Assert.Equal("C1", local.Range);
        Assert.Equal("Hoja1", local.Scope!.Name);

        ExcelDefinedName constant = workbook.FindDefinedName("Constante")!;
        Assert.NotNull(constant);
        Assert.False(constant.IsRange);
        Assert.Null(constant.Range);
        Assert.Equal("42", constant.RefersTo);

        Assert.NotNull(workbook.FindDefinedName("_xlnm.Print_Area"));
        Assert.Null(workbook.FindDefinedName("NoExiste"));
    }

    [Fact]
    public void DefinedNames_ArePreservedWhenTheWorkbookIsSaved()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithDefinedNames(
            ("Datos", "Hoja1!$A$1:$B$2", null),
            ("Constante", "42", null));

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            produced = SaveBytes.ToBytes(workbook);
        }

        IReadOnlyList<(string Name, string Text, uint? LocalSheetId)> names =
            OoxmlValidation.ReadXlsxDefinedNames(produced);
        Assert.Equal(2, names.Count);
        Assert.Equal("Hoja1!$A$1:$B$2", names[0].Text);
        Assert.Equal("42", names[1].Text);
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void RemoveSheet_MovesTheScopeOfTheRemainingLocalNames()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet first = workbook.Worksheets[0];
            Worksheet second = workbook.Worksheets.Add("Otra");
            second.AddLocalNamedRange("Local", "A1");

            Assert.True(workbook.Worksheets.Remove("Sheet1"));

            Assert.Single(workbook.DefinedNames);
            Assert.Equal("Otra", workbook.DefinedNames[0].Scope!.Name);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        IReadOnlyList<(string Name, string Text, uint? LocalSheetId)> names =
            OoxmlValidation.ReadXlsxDefinedNames(bytes);
        Assert.Single(names);
        Assert.Equal(0U, names[0].LocalSheetId);
    }

    [Fact]
    public void DefinedNames_ThatAreNotRanges_AreListedButNotInterpreted()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithDefinedNames(
            ("Desconocida", "Otra!$A$1", null),
            ("Malformada", "$A$1:", null),
            ("Vacia", "", null),
            ("Formula", "=Hoja1!$A$1+1", null));

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        Assert.Equal(4, workbook.DefinedNames.Count);
        foreach (ExcelDefinedName definedName in workbook.DefinedNames)
        {
            Assert.False(definedName.IsRange);
            Assert.Null(definedName.Range);
        }
    }

    [Fact]
    public void DisposedWorkbook_DefinedNamesThrowsObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = workbook.DefinedNames);
        Assert.Throws<ObjectDisposedException>(() => _ = workbook.FindDefinedName("Datos"));
        Assert.Throws<ObjectDisposedException>(() => workbook.RemoveNamedRange("Datos"));
    }
}
