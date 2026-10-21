using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.8: interoperability of inline images and relationships, verified with the Open XML SDK
/// instead of the code under test.
/// </summary>
public sealed class WordImageInteropTests
{
    [Fact]
    public void ProducedDocument_WithImage_IsSchemaValidAndRelationshipResolves()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        Add(paragraph, 36.0, 18.0);

        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(1, OoxmlValidation.CountDocxImageParts(produced));
        Assert.Equal(["image/png"], OoxmlValidation.ReadDocxImageContentTypes(produced));
        Assert.Equal(OoxmlValidation.TinyPngBytes, Assert.Single(OoxmlValidation.ReadDocxImageBytes(produced)));
        Assert.Equal([(457200L, 228600L)], OoxmlValidation.ReadDocxAllInlineExtents(produced));

        using MemoryStream package = new(produced, writable: false);
        using WordprocessingDocument sdk = WordprocessingDocument.Open(package, isEditable: false);
        MainDocumentPart mainPart = sdk.MainDocumentPart!;
        A.Blip blip = Assert.Single(mainPart.Document!.Descendants<A.Blip>());
        Assert.NotNull(blip.Embed);
        Assert.True(mainPart.TryGetPartById(blip.Embed!, out OpenXmlPart? related));
        ImagePart imagePart = Assert.IsType<ImagePart>(related);
        Assert.Equal("image/png", imagePart.ContentType);
    }

    [Fact]
    public void OpenedDocument_WithInlineImage_PreservesItWhenTheBodyIsEdited()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithInlineImage("antes", 12700, 12700);
        byte[] produced;

        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Assert.Single(document.Paragraphs[1].Images);
            document.AddParagraph("después");
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(1, OoxmlValidation.CountDocxImageParts(produced));
        Assert.Equal(OoxmlValidation.TinyPngBytes, Assert.Single(OoxmlValidation.ReadDocxImageBytes(produced)));
        Assert.Equal([1U], OoxmlValidation.ReadDocxAllDrawingIds(produced));

        using WordDocument check = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Single(check.Paragraphs[1].Images);
        Assert.Equal("después", check.Paragraphs[2].Text);
    }

    [Fact]
    public void OpenedDocument_WithHyperlinkRelationship_PreservesItWhenAnImageIsAdded()
    {
        const string uri = "https://example.com/engineoffice";
        byte[] fixture = OoxmlValidation.CreateDocxWithHyperlink("enlace", uri);
        string relationshipId = Assert.Single(OoxmlValidation.ReadDocxHyperlinkReferenceIds(fixture));
        byte[] produced;

        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Add(document.AddParagraph(string.Empty), 10.0, 10.0);
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal([uri], OoxmlValidation.ReadDocxExternalHyperlinkUris(produced));
        Assert.Equal([relationshipId], OoxmlValidation.ReadDocxHyperlinkReferenceIds(produced));
        Assert.Equal(1, OoxmlValidation.CountDocxImageParts(produced));
    }

    [Fact]
    public void ProducedDocument_WithImagesInBodyTableHeaderAndFooter_IsSchemaValid()
    {
        using WordDocument document = WordDocument.Create();

        Add(document.AddParagraph(string.Empty), 10.0, 10.0);
        Add(document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0], 11.0, 11.0);
        Add(document.Sections[0].AddHeader(HeaderFooterType.Primary).Paragraphs[0], 12.0, 12.0);
        Add(document.Sections[0].AddFooter(HeaderFooterType.Primary).Paragraphs[0], 13.0, 13.0);

        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(2, OoxmlValidation.CountDocxImageParts(produced));
        Assert.Equal(1, OoxmlValidation.CountDocxHeaderImageParts(produced));
        Assert.Equal(1, OoxmlValidation.CountDocxFooterImageParts(produced));
        Assert.Equal(4, OoxmlValidation.ReadDocxAllDrawingIds(produced).Count);
        Assert.Equal(4, OoxmlValidation.ReadDocxAllDrawingIds(produced).Distinct().Count());

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Single(reopened.Paragraphs[0].Images);
        Assert.Single(reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Images);
        Assert.Single(reopened.Sections[0].Header(HeaderFooterType.Primary)!.Paragraphs[0].Images);
        Assert.Single(reopened.Sections[0].Footer(HeaderFooterType.Primary)!.Paragraphs[0].Images);
    }

    private static void Add(Paragraph paragraph, double widthPoints, double heightPoints)
    {
        using MemoryStream stream = new(OoxmlValidation.TinyPngBytes, writable: false);
        paragraph.AddImage(stream, ImageFormat.Png, widthPoints, heightPoints);
    }
}
