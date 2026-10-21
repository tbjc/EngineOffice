using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.9: interoperability of external hyperlinks, verified with the Open XML SDK instead of
/// the code under test.
/// </summary>
public sealed class WordHyperlinkInteropTests
{
    [Fact]
    public void ProducedDocument_WithHyperlink_IsSchemaValidAndRelationshipResolves()
    {
        Uri destination = new("https://example.com/hyperlink");
        byte[] produced;

        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("antes ").AddHyperlink("ejemplo", destination);
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal("antes ejemplo", OoxmlValidation.ReadDocxText(produced));
        Assert.Equal(["ejemplo"], OoxmlValidation.ReadDocxHyperlinkTexts(produced));
        Assert.Equal([destination.OriginalString], OoxmlValidation.ReadDocxExternalHyperlinkUris(produced));

        using MemoryStream package = new(produced, writable: false);
        using WordprocessingDocument sdk = WordprocessingDocument.Open(package, isEditable: false);
        MainDocumentPart mainPart = sdk.MainDocumentPart!;
        W.Hyperlink stored = Assert.Single(mainPart.Document!.Descendants<W.Hyperlink>());
        Assert.Equal("ejemplo", stored.InnerText);

        HyperlinkRelationship relationship = Assert.Single(mainPart.HyperlinkRelationships);
        Assert.Equal(stored.Id?.Value, relationship.Id);
        Assert.True(relationship.IsExternal);
        Assert.Equal(destination, relationship.Uri);
    }

    [Fact]
    public void OpenedDocument_WithHyperlink_CanEditTextAndDestination()
    {
        Uri oldDestination = new("https://example.com/old");
        Uri newDestination = new("https://example.com/new");
        byte[] fixture = OoxmlValidation.CreateDocxWithHyperlink("viejo", oldDestination.OriginalString);
        byte[] produced;

        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Hyperlink link = Assert.Single(document.Paragraphs[0].Hyperlinks);
            link.Text = "nuevo";
            link.Destination = newDestination;
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(["nuevo"], OoxmlValidation.ReadDocxHyperlinkTexts(produced));
        Assert.Equal([newDestination.OriginalString], OoxmlValidation.ReadDocxExternalHyperlinkUris(produced));
        Assert.Equal([newDestination.OriginalString], OoxmlValidation.ReadDocxHyperlinkDestinations(produced));

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Hyperlink reopenedLink = Assert.Single(reopened.Paragraphs[0].Hyperlinks);
        Assert.Equal("nuevo", reopenedLink.Text);
        Assert.Equal(newDestination, reopenedLink.Destination);
    }

    [Fact]
    public void OpenedDocument_WithInternalHyperlink_PreservesItWhenTheBodyIsEdited()
    {
        const string anchor = "seccion";
        byte[] fixture = OoxmlValidation.CreateDocxWithInternalHyperlink("interno", anchor);
        byte[] produced;

        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Assert.Empty(document.Paragraphs[0].Hyperlinks);
            document.AddParagraph("después");
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal([anchor], OoxmlValidation.ReadDocxHyperlinkAnchors(produced));
        Assert.Empty(OoxmlValidation.ReadDocxExternalHyperlinkUris(produced));
        Assert.Equal("internodespués", OoxmlValidation.ReadDocxText(produced));
    }

    [Fact]
    public void ProducedDocument_WithHyperlinksInBodyTableHeaderAndFooter_IsSchemaValid()
    {
        Uri bodyDestination = new("https://example.com/body");
        Uri cellDestination = new("https://example.com/cell");
        Uri headerDestination = new("https://example.com/header");
        Uri footerDestination = new("https://example.com/footer");
        byte[] produced;

        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph(string.Empty).AddHyperlink("cuerpo", bodyDestination);
            document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0]
                .AddHyperlink("celda", cellDestination);
            document.Sections[0].AddHeader(HeaderFooterType.Primary).Paragraphs[0]
                .AddHyperlink("cabecera", headerDestination);
            document.Sections[0].AddFooter(HeaderFooterType.Primary).Paragraphs[0]
                .AddHyperlink("pie", footerDestination);
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(2, OoxmlValidation.ReadDocxExternalHyperlinkUris(produced).Count);
        Assert.Equal([headerDestination.OriginalString],
            OoxmlValidation.ReadDocxHeaderExternalHyperlinkUris(produced));
        Assert.Equal([footerDestination.OriginalString],
            OoxmlValidation.ReadDocxFooterExternalHyperlinkUris(produced));

        using MemoryStream package = new(produced, writable: false);
        using WordprocessingDocument sdk = WordprocessingDocument.Open(package, isEditable: false);
        MainDocumentPart mainPart = sdk.MainDocumentPart!;
        W.Hyperlink[] bodyLinks = mainPart.Document!.Descendants<W.Hyperlink>().ToArray();
        Assert.Equal(2, bodyLinks.Length);
        Assert.All(bodyLinks, link =>
        {
            Assert.NotNull(link.Id);
            Assert.Contains(mainPart.HyperlinkRelationships, relationship => relationship.Id == link.Id!.Value);
        });

        HeaderPart headerPart = Assert.Single(mainPart.HeaderParts);
        W.Hyperlink headerLink = Assert.Single(headerPart.Header!.Descendants<W.Hyperlink>());
        HyperlinkRelationship headerRelationship = Assert.Single(headerPart.HyperlinkRelationships);
        Assert.Equal(headerLink.Id?.Value, headerRelationship.Id);
        Assert.Equal(headerDestination, headerRelationship.Uri);

        FooterPart footerPart = Assert.Single(mainPart.FooterParts);
        W.Hyperlink footerLink = Assert.Single(footerPart.Footer!.Descendants<W.Hyperlink>());
        HyperlinkRelationship footerRelationship = Assert.Single(footerPart.HyperlinkRelationships);
        Assert.Equal(footerLink.Id?.Value, footerRelationship.Id);
        Assert.Equal(footerDestination, footerRelationship.Uri);
    }

    [Fact]
    public void OpenedDocument_WithTwoHyperlinksSharingARelationship_KeepsTheRelationshipInUse()
    {
        Uri oldDestination = new("https://example.com/shared");
        Uri newDestination = new("https://example.com/edited");
        byte[] fixture = OoxmlValidation.CreateDocxWithTwoHyperlinksSharingRelationship(
            "primero",
            "segundo",
            oldDestination.OriginalString);
        byte[] produced;

        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            document.Paragraphs[0].Hyperlinks[0].Destination = newDestination;
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(
            [newDestination.OriginalString, oldDestination.OriginalString],
            OoxmlValidation.ReadDocxHyperlinkDestinations(produced));
        Assert.Equal(2, OoxmlValidation.ReadDocxExternalHyperlinkUris(produced).Count);
    }

    [Fact]
    public void OpenedDocument_WithHyperlinkInAHeader_EditsTheRelationshipOfTheHeaderPart()
    {
        Uri oldDestination = new("https://example.com/header-old");
        Uri newDestination = new("https://example.com/header-new");
        byte[] fixture;
        byte[] produced;

        using (WordDocument document = WordDocument.Create())
        {
            document.Sections[0].AddHeader(HeaderFooterType.Primary).Paragraphs[0]
                .AddHyperlink("cabecera", oldDestination);
            fixture = SaveBytes.ToBytes(document);
        }

        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Hyperlink link = Assert.Single(
                document.Sections[0].Header(HeaderFooterType.Primary)!.Paragraphs[0].Hyperlinks);
            link.Destination = newDestination;
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Empty(OoxmlValidation.ReadDocxExternalHyperlinkUris(produced));
        Assert.Equal([newDestination.OriginalString],
            OoxmlValidation.ReadDocxHeaderExternalHyperlinkUris(produced));

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Equal(newDestination, Assert.Single(
            reopened.Sections[0].Header(HeaderFooterType.Primary)!.Paragraphs[0].Hyperlinks).Destination);
    }
}
