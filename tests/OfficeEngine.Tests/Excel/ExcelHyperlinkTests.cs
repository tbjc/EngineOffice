using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>Hyperlinks of a worksheet: external addresses and internal locations.</summary>
public sealed class ExcelHyperlinkTests
{
    [Fact]
    public void AddHyperlink_WritesTheAddressAndTheRelationship()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            ExcelHyperlink hyperlink = sheet.AddHyperlink(
                "A1",
                "https://example.com/docs",
                display: "Documentación",
                tooltip: "Abrir la documentación");

            Assert.Equal("A1", hyperlink.Reference);
            Assert.Equal("https://example.com/docs", hyperlink.Target);
            Assert.True(hyperlink.IsExternal);
            Assert.Null(hyperlink.Location);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        IReadOnlyList<(string Reference, string? Target, string? Location)> links =
            OoxmlValidation.ReadXlsxHyperlinks(bytes, "Sheet1");
        Assert.Single(links);
        Assert.Equal("A1", links[0].Reference);
        Assert.Equal("https://example.com/docs", links[0].Target);
        Assert.Null(links[0].Location);

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        ExcelHyperlink read = reopened.Worksheets[0].Hyperlinks[0];
        Assert.Equal("https://example.com/docs", read.Target);
        Assert.Equal("Documentación", read.Display);
        Assert.Equal("Abrir la documentación", read.Tooltip);
    }

    [Fact]
    public void AddInternalHyperlink_WritesTheLocation()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            ExcelHyperlink hyperlink = sheet.AddInternalHyperlink("B2", "Resumen!A1");

            Assert.Equal("Resumen!A1", hyperlink.Location);
            Assert.Null(hyperlink.Target);
            Assert.False(hyperlink.IsExternal);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        IReadOnlyList<(string Reference, string? Target, string? Location)> links =
            OoxmlValidation.ReadXlsxHyperlinks(bytes, "Sheet1");
        Assert.Single(links);
        Assert.Null(links[0].Target);
        Assert.Equal("Resumen!A1", links[0].Location);
    }

    [Fact]
    public void AddHyperlink_WithInvalidValues_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => sheet.AddHyperlink("B2:A1", "https://example.com"));
        Assert.Throws<ArgumentException>(() => sheet.AddHyperlink("A1", string.Empty));
        Assert.Throws<ArgumentException>(() => sheet.AddHyperlink("A1", "no-es-una-direccion"));
        Assert.Throws<ArgumentException>(() => sheet.AddInternalHyperlink("A1", string.Empty));
        Assert.Empty(sheet.Hyperlinks);
    }

    [Fact]
    public void RemoveHyperlink_RemovesTheLinkAndTheRelationship()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.AddHyperlink("A1", "https://example.com");
            sheet.AddInternalHyperlink("B2", "Resumen!A1");

            Assert.True(sheet.RemoveHyperlink("A1"));
            Assert.False(sheet.RemoveHyperlink("A1"));
            Assert.Single(sheet.Hyperlinks);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        IReadOnlyList<(string Reference, string? Target, string? Location)> links =
            OoxmlValidation.ReadXlsxHyperlinks(bytes, "Sheet1");
        Assert.Single(links);
        Assert.Equal("B2", links[0].Reference);
    }

    [Fact]
    public void Hyperlinks_ReadsTheLinksOfAnotherWorkbook()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithHyperlinks(
            ("A1", "https://example.com", null),
            ("B2", null, "Resumen!A1"));

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        Assert.Equal(2, workbook.Worksheets[0].Hyperlinks.Count);
        Assert.Equal("https://example.com/", workbook.Worksheets[0].Hyperlinks[0].Target);
        Assert.Equal("Resumen!A1", workbook.Worksheets[0].Hyperlinks[1].Location);
    }

    [Fact]
    public void Hyperlinks_ArePreservedWhenTheWorkbookIsSaved()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithHyperlinks(("A1", "https://example.com", null));

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Single(OoxmlValidation.ReadXlsxHyperlinks(produced, "Hoja1"));
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void DisposedWorkbook_HyperlinksThrowObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].Hyperlinks);
    }
}
