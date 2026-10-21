using System;
using System.IO;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.8: inline images and the relationships that carry them. A created image must become a
/// package part with a resolvable relationship, and an image read from a foreign document must keep
/// its bytes and dimensions.
/// </summary>
public sealed class WordImageTests
{
    [Fact]
    public void Create_HasNoImages()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);

        Assert.Empty(paragraph.Images);
    }

    [Fact]
    public void AddImage_FromFile_RoundTripsThroughAFile()
    {
        string directory = TestData.NewTemporaryDirectory();
        string path = Path.Combine(directory, "pixel.png");
        string documentPath = Path.Combine(directory, "image.docx");
        try
        {
            File.WriteAllBytes(path, OoxmlValidation.TinyPngBytes);

            using (WordDocument document = WordDocument.Create())
            {
                Paragraph paragraph = document.AddParagraph("Antes");
                Image image = paragraph.AddImage(path, 36.0, 18.0);

                Assert.Equal("image/png", image.ContentType);
                Assert.Equal(36.0, image.WidthPoints);
                Assert.Equal(18.0, image.HeightPoints);
                Assert.Equal(OoxmlValidation.TinyPngBytes, image.ReadAllBytes());
                Assert.Single(paragraph.Images);
                Assert.Equal("Antes", paragraph.Text);

                document.Save(documentPath);
            }

            using (WordDocument reopened = WordDocument.Open(documentPath))
            {
                Image image = Assert.Single(reopened.Paragraphs[0].Images);
                Assert.Equal("image/png", image.ContentType);
                Assert.Equal(36.0, image.WidthPoints);
                Assert.Equal(18.0, image.HeightPoints);
                Assert.Equal(OoxmlValidation.TinyPngBytes, image.ReadAllBytes());
            }
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void AddImage_FromStream_ConvertsPointsToEmus()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        using MemoryStream stream = new(OoxmlValidation.TinyPngBytes, writable: false);

        Image image = paragraph.AddImage(stream, ImageFormat.Png, 12.5, 7.25);

        Assert.Equal(12.5, image.WidthPoints);
        Assert.Equal(7.25, image.HeightPoints);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal([(158750L, 92075L)], OoxmlValidation.ReadDocxAllInlineExtents(produced));
        Assert.Equal(OoxmlValidation.TinyPngBytes, Assert.Single(OoxmlValidation.ReadDocxImageBytes(produced)));
    }

    [Theory]
    [InlineData(ImageFormat.Png, "image/png")]
    [InlineData(ImageFormat.Jpeg, "image/jpeg")]
    [InlineData(ImageFormat.Gif, "image/gif")]
    [InlineData(ImageFormat.Bmp, "image/bmp")]
    [InlineData(ImageFormat.Tiff, "image/tiff")]
    [InlineData(ImageFormat.Emf, "image/x-emf")]
    [InlineData(ImageFormat.Wmf, "image/x-wmf")]
    public void AddImage_WithEachSupportedFormat_WritesThatContentType(ImageFormat format, string contentType)
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        using MemoryStream stream = new(OoxmlValidation.TinyPngBytes, writable: false);

        Image image = paragraph.AddImage(stream, format, 1.0, 1.0);

        Assert.Equal(contentType, image.ContentType);
        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal([contentType], OoxmlValidation.ReadDocxImageContentTypes(produced));
    }

    [Fact]
    public void AddImage_KeepsExistingContentAndAddsARun()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("Antes");
        using MemoryStream stream = new(OoxmlValidation.TinyPngBytes, writable: false);

        paragraph.AddImage(stream, ImageFormat.Png, 10.0, 10.0);

        Assert.Equal("Antes", paragraph.Text);
        Assert.Equal(2, paragraph.Runs.Count);
        Assert.Single(paragraph.Images);
    }

    [Fact]
    public void AddImage_InATableCell_RoundTrips()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(1, 1);
        Paragraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        using MemoryStream stream = new(OoxmlValidation.TinyPngBytes, writable: false);

        paragraph.AddImage(stream, ImageFormat.Png, 8.0, 8.0);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(1, OoxmlValidation.CountDocxImageParts(produced));

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Single(reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Images);
    }

    [Fact]
    public void AddImage_InAHeader_RoundTrips()
    {
        using WordDocument document = WordDocument.Create();
        HeaderFooter header = document.Sections[0].AddHeader(HeaderFooterType.Primary);
        using MemoryStream stream = new(OoxmlValidation.TinyPngBytes, writable: false);

        header.Paragraphs[0].AddImage(stream, ImageFormat.Png, 16.0, 9.0);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(1, OoxmlValidation.CountDocxHeaderImageParts(produced));

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Image image = Assert.Single(reopened.Sections[0].Header(HeaderFooterType.Primary)!.Paragraphs[0].Images);
        Assert.Equal(16.0, image.WidthPoints);
        Assert.Equal(9.0, image.HeightPoints);
    }

    [Fact]
    public void AddImage_Twice_UsesDistinctDrawingIds()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        using MemoryStream first = new(OoxmlValidation.TinyPngBytes, writable: false);
        using MemoryStream second = new(OoxmlValidation.TinyPngBytes, writable: false);

        paragraph.AddImage(first, ImageFormat.Png, 10.0, 10.0);
        paragraph.AddImage(second, ImageFormat.Png, 20.0, 20.0);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(2, OoxmlValidation.CountDocxImageParts(produced));
        Assert.Equal(2, OoxmlValidation.ReadDocxAllDrawingIds(produced).Distinct().Count());
        Assert.Equal(2, paragraph.Images.Count);
    }

    [Fact]
    public void AddImage_RejectsInvalidArguments()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);

        Assert.Throws<ArgumentNullException>(() => paragraph.AddImage((string)null!, 1.0, 1.0));
        Assert.Throws<ArgumentException>(() => paragraph.AddImage(string.Empty, 1.0, 1.0));
        Assert.Throws<ArgumentException>(() => paragraph.AddImage("picture.txt", 1.0, 1.0));
        Assert.Throws<ArgumentNullException>(() => paragraph.AddImage((Stream)null!, ImageFormat.Png, 1.0, 1.0));
        Assert.Throws<ArgumentOutOfRangeException>(
            () => paragraph.AddImage(new MemoryStream(OoxmlValidation.TinyPngBytes), (ImageFormat)99, 1.0, 1.0));

        using MemoryStream stream = new(OoxmlValidation.TinyPngBytes, writable: false);
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.AddImage(stream, ImageFormat.Png, 0.0, 1.0));
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.AddImage(stream, ImageFormat.Png, 1.0, -1.0));
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.AddImage(stream, ImageFormat.Png, double.NaN, 1.0));
        Assert.Throws<ArgumentOutOfRangeException>(
            () => paragraph.AddImage(stream, ImageFormat.Png, 1.0, double.PositiveInfinity));
    }

    [Fact]
    public void AddImage_RejectsAnUnreadableStream()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        MemoryStream stream = new();
        stream.Dispose();

        Assert.Throws<ArgumentException>(() => paragraph.AddImage(stream, ImageFormat.Png, 1.0, 1.0));
    }

    [Fact]
    public void AddImage_WithInvalidDimensions_LeavesNoImagePart()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        using MemoryStream stream = new(OoxmlValidation.TinyPngBytes, writable: false);

        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.AddImage(stream, ImageFormat.Png, 0.0, 1.0));

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(paragraph.Images);
        Assert.Equal(0, OoxmlValidation.CountDocxImageParts(produced));
    }

    [Fact]
    public void Images_ReadsAnExistingInlineImage()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithInlineImage("cuerpo", widthEmus: 25400, heightEmus: 12700);

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Image image = Assert.Single(document.Paragraphs[1].Images);

        Assert.Equal(2.0, image.WidthPoints);
        Assert.Equal(1.0, image.HeightPoints);
        Assert.Equal("image/png", image.ContentType);
        Assert.Equal(OoxmlValidation.TinyPngBytes, image.ReadAllBytes());
        Assert.NotEmpty(image.RelationshipId);
    }

    [Fact]
    public void Images_ReportsAMissingImagePart()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithInlineImageReferenceToMissingPart("cuerpo");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[1].Images);
    }

    [Fact]
    public void Images_ReportsAnInlineImageWithoutExtent()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithInlineImageWithoutExtent("cuerpo");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[1].Images);
    }

    [Fact]
    public void Images_ReportsAnImageRelationshipToANonImagePart()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithInlineImagePointingToANonImagePart("cuerpo");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[1].Images);
    }
}
