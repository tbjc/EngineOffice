using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>Images anchored in a worksheet: the media part, the drawing and its relationship.</summary>
public sealed class ExcelImageTests
{
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    [Fact]
    public void AddImage_WritesTheMediaPartAndTheDrawing()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            ExcelImage image = sheet.AddImage(
                "B2",
                new MemoryStream(PngBytes),
                ExcelImageFormat.Png,
                100d,
                50d);

            Assert.Equal("B2", image.Reference);
            Assert.Equal(100d, image.WidthPoints);
            Assert.Equal(50d, image.HeightPoints);
            Assert.Equal("image/png", image.MediaType);
            Assert.Single(sheet.Images);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.Contains(TestPackage.EntryNames(bytes), name => name.StartsWith("xl/media/", StringComparison.Ordinal));
        Assert.True(TestPackage.HasEntry(bytes, "xl/drawings/drawing1.xml"));

        string sheetXml = TestPackage.ReadEntryText(bytes, "xl/worksheets/sheet1.xml");
        Assert.Contains("<x:drawing", sheetXml, StringComparison.Ordinal);

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        ExcelImage read = reopened.Worksheets[0].Images[0];
        Assert.Equal("B2", read.Reference);
        Assert.Equal(100d, read.WidthPoints);
        Assert.Equal("image/png", read.MediaType);
    }

    [Fact]
    public void AddImage_WithSeveralImages_KeepsThemAll()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.AddImage("A1", new MemoryStream(PngBytes), ExcelImageFormat.Png, 10d, 10d, "Uno");
            sheet.AddImage("C3", new MemoryStream(PngBytes), ExcelImageFormat.Png, 20d, 20d, "Dos");

            Assert.Equal(2, sheet.Images.Count);
            Assert.Equal("A1", sheet.Images[0].Reference);
            Assert.Equal("C3", sheet.Images[1].Reference);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void AddImage_WithInvalidValues_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => sheet.AddImage(
            "B2:A1",
            new MemoryStream(PngBytes),
            ExcelImageFormat.Png,
            10d,
            10d));
        Assert.Throws<ArgumentException>(() => sheet.AddImage(
            "A1",
            new MemoryStream(),
            ExcelImageFormat.Png,
            10d,
            10d));
        Assert.Throws<ArgumentException>(() => sheet.AddImage(
            "A1",
            new MemoryStream(PngBytes),
            ExcelImageFormat.Png,
            0d,
            10d));
        Assert.Empty(sheet.Images);
    }

    [Fact]
    public void RemoveImage_RemovesTheAnchor()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.AddImage("A1", new MemoryStream(PngBytes), ExcelImageFormat.Png, 10d, 10d, "Imagen");

            Assert.True(sheet.RemoveImage("Imagen"));
            Assert.False(sheet.RemoveImage("Imagen"));
            Assert.Empty(sheet.Images);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void Images_ReadsTheImagesOfAnotherWorkbook()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithImage("Logo", "A1", PngBytes);

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        ExcelImage image = workbook.Worksheets[0].Images[0];
        Assert.Equal("Logo", image.Name);
        Assert.Equal("A1", image.Reference);
        Assert.Equal(100d, image.WidthPoints);
        Assert.Equal(50d, image.HeightPoints);
        Assert.Equal("image/png", image.MediaType);
    }

    [Fact]
    public void Images_ArePreservedWhenTheWorkbookIsSaved()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithImage("Logo", "A1", PngBytes);

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Contains(TestPackage.EntryNames(produced), name => name.StartsWith("xl/media/", StringComparison.Ordinal));
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void DisposedWorkbook_ImagesThrowObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].Images);
    }
}
