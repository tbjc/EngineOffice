using System;
using System.IO;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.9: external hyperlinks. A link must keep its visible text and destination through a
/// round-trip, and both must remain editable without breaking relationships this version does not
/// model.
/// </summary>
public sealed class WordHyperlinkTests
{
    [Fact]
    public void Create_HasNoHyperlinks()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);

        Assert.Empty(paragraph.Hyperlinks);
    }

    [Fact]
    public void AddHyperlink_RoundTripsThroughAFile()
    {
        string directory = TestData.NewTemporaryDirectory();
        string documentPath = Path.Combine(directory, "hyperlink.docx");
        Uri destination = new("https://example.com/engineoffice");
        try
        {
            using (WordDocument document = WordDocument.Create())
            {
                Paragraph paragraph = document.AddParagraph("Antes ");
                Hyperlink link = paragraph.AddHyperlink("Motor", destination);

                Assert.Equal("Motor", link.Text);
                Assert.Equal(destination, link.Destination);
                Assert.Equal("Antes Motor", paragraph.Text);
                Assert.Single(paragraph.Hyperlinks);
                Assert.Single(paragraph.Runs);
                Assert.Equal("Motor", Assert.Single(link.Runs).Text);

                document.Save(documentPath);
            }

            using (WordDocument reopened = WordDocument.Open(documentPath))
            {
                Paragraph paragraph = reopened.Paragraphs[0];
                Hyperlink link = Assert.Single(paragraph.Hyperlinks);
                Assert.Equal("Antes Motor", paragraph.Text);
                Assert.Equal("Motor", link.Text);
                Assert.Equal(destination, link.Destination);
            }
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void AddHyperlink_RoundTripsThroughAStream()
    {
        Uri destination = new("https://example.com/stream");
        byte[] produced;

        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph(string.Empty).AddHyperlink("stream", destination);
            produced = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Hyperlink link = Assert.Single(reopened.Paragraphs[0].Hyperlinks);
        Assert.Equal("stream", link.Text);
        Assert.Equal(destination, link.Destination);
    }

    [Fact]
    public void AddHyperlink_InATableCell_RoundTrips()
    {
        Uri destination = new("https://example.com/cell");
        byte[] produced;

        using (WordDocument document = WordDocument.Create())
        {
            Paragraph paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
            paragraph.AddHyperlink("celda", destination);
            produced = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Hyperlink link = Assert.Single(reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Hyperlinks);
        Assert.Equal("celda", link.Text);
        Assert.Equal(destination, link.Destination);
    }

    [Fact]
    public void AddHyperlink_InAHeaderAndFooter_RoundTrips()
    {
        Uri headerDestination = new("https://example.com/header");
        Uri footerDestination = new("https://example.com/footer");
        byte[] produced;

        using (WordDocument document = WordDocument.Create())
        {
            document.Sections[0].AddHeader(HeaderFooterType.Primary).Paragraphs[0]
                .AddHyperlink("cabecera", headerDestination);
            document.Sections[0].AddFooter(HeaderFooterType.Primary).Paragraphs[0]
                .AddHyperlink("pie", footerDestination);
            produced = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Hyperlink header = Assert.Single(
            reopened.Sections[0].Header(HeaderFooterType.Primary)!.Paragraphs[0].Hyperlinks);
        Hyperlink footer = Assert.Single(
            reopened.Sections[0].Footer(HeaderFooterType.Primary)!.Paragraphs[0].Hyperlinks);

        Assert.Equal(headerDestination, header.Destination);
        Assert.Equal(footerDestination, footer.Destination);
    }

    [Fact]
    public void AddHyperlink_Twice_CreatesTwoRelationships()
    {
        Uri destination = new("https://example.com/same");
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);

        paragraph.AddHyperlink("uno", destination);
        paragraph.AddHyperlink("dos", destination);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Equal(2, paragraph.Hyperlinks.Count);
        Assert.Equal([destination.OriginalString, destination.OriginalString],
            OoxmlValidation.ReadDocxExternalHyperlinkUris(produced));
    }

    [Fact]
    public void Hyperlink_Text_IsEditableAndKeepsItsDestination()
    {
        Uri destination = new("https://example.com/text");
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        Hyperlink link = paragraph.AddHyperlink("viejo", destination);

        link.Text = "nuevo";

        Assert.Equal("nuevo", link.Text);
        Assert.Equal(destination, link.Destination);
        Assert.Equal("nuevo", paragraph.Text);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Equal(["nuevo"], OoxmlValidation.ReadDocxHyperlinkTexts(produced));
        Assert.Equal([destination.OriginalString], OoxmlValidation.ReadDocxExternalHyperlinkUris(produced));
    }

    [Fact]
    public void Hyperlink_TextSetter_KeepsTheRelationshipAndDropsTheOldRuns()
    {
        Uri destination = new("https://example.com/run");
        using WordDocument document = WordDocument.Create();
        Hyperlink link = document.AddParagraph(string.Empty).AddHyperlink("uno", destination);
        link.Runs[0].Bold = true;

        link.Text = "dos";

        Assert.Equal("dos", link.Text);
        Assert.Null(Assert.Single(link.Runs).Bold);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Equal([destination.OriginalString], OoxmlValidation.ReadDocxExternalHyperlinkUris(produced));
    }

    [Fact]
    public void Hyperlink_Runs_EditTextWithoutLosingFormatting()
    {
        Uri destination = new("https://example.com/format");
        using WordDocument document = WordDocument.Create();
        Hyperlink link = document.AddParagraph(string.Empty).AddHyperlink("uno", destination);
        link.Runs[0].Bold = true;

        link.Runs[0].Text = "dos";

        Assert.Equal("dos", link.Text);
        Assert.True(link.Runs[0].Bold);

        byte[] produced = SaveBytes.ToBytes(document);
        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Run run = Assert.Single(Assert.Single(reopened.Paragraphs[0].Hyperlinks).Runs);
        Assert.Equal("dos", run.Text);
        Assert.True(run.Bold);
    }

    [Fact]
    public void Hyperlink_Destination_IsEditableAndRemovesTheOldRelationship()
    {
        Uri oldDestination = new("https://example.com/old");
        Uri newDestination = new("https://example.com/new");
        using WordDocument document = WordDocument.Create();
        Hyperlink link = document.AddParagraph("texto").AddHyperlink("enlace", oldDestination);

        link.Destination = newDestination;

        Assert.Equal(newDestination, link.Destination);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Equal([newDestination.OriginalString], OoxmlValidation.ReadDocxHyperlinkDestinations(produced));
        Assert.Equal([newDestination.OriginalString], OoxmlValidation.ReadDocxExternalHyperlinkUris(produced));

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Equal(newDestination, Assert.Single(reopened.Paragraphs[0].Hyperlinks).Destination);
    }

    [Fact]
    public void Hyperlink_Destination_WhenTheRelationshipIsShared_LeavesTheOtherLinkAlone()
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
            Hyperlink first = document.Paragraphs[0].Hyperlinks[0];
            first.Destination = newDestination;
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Equal(
            [newDestination.OriginalString, oldDestination.OriginalString],
            OoxmlValidation.ReadDocxHyperlinkDestinations(produced));
        Assert.Equal(2, OoxmlValidation.ReadDocxExternalHyperlinkUris(produced).Count);

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Equal(newDestination, reopened.Paragraphs[0].Hyperlinks[0].Destination);
        Assert.Equal(oldDestination, reopened.Paragraphs[0].Hyperlinks[1].Destination);
    }

    [Fact]
    public void Hyperlinks_ReadsAnExistingExternalHyperlink()
    {
        Uri destination = new("https://example.com/foreign");
        byte[] fixture = OoxmlValidation.CreateDocxWithHyperlink("ajeno", destination.OriginalString);

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Hyperlink link = Assert.Single(document.Paragraphs[0].Hyperlinks);

        Assert.Equal("ajeno", link.Text);
        Assert.Equal(destination, link.Destination);
        Assert.Equal("ajeno", Assert.Single(link.Runs).Text);
    }

    [Fact]
    public void Hyperlinks_DoesNotReturnAnInternalAnchorLink()
    {
        const string anchor = "destino";
        byte[] fixture = OoxmlValidation.CreateDocxWithInternalHyperlink("interno", anchor);
        byte[] produced;

        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Assert.Empty(document.Paragraphs[0].Hyperlinks);
            document.AddParagraph("después");
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Equal([anchor], OoxmlValidation.ReadDocxHyperlinkAnchors(produced));
        Assert.Empty(OoxmlValidation.ReadDocxExternalHyperlinkUris(produced));
    }

    [Fact]
    public void Hyperlinks_ReportsAMissingRelationship()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithHyperlinkReferenceToMissingRelationship("roto");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].Hyperlinks);
    }

    [Fact]
    public void Hyperlinks_ReportsARelationshipThatIsNotAHyperlink()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithHyperlinkReferenceToNonHyperlinkPart("roto");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].Hyperlinks);
    }

    [Fact]
    public void Hyperlinks_ReportsANonExternalHyperlinkRelationship()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithNonExternalHyperlinkRelationship("roto");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].Hyperlinks);
    }

    [Fact]
    public void AddHyperlink_RejectsInvalidArgumentsWithoutCreatingARelationship()
    {
        Uri destination = new("https://example.com/valid");
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);

        Assert.Throws<ArgumentNullException>(() => paragraph.AddHyperlink(null!, destination));
        Assert.Throws<ArgumentNullException>(() => paragraph.AddHyperlink("texto", null!));
        Assert.Throws<ArgumentException>(() => paragraph.AddHyperlink("a\tb", destination));
        Assert.Throws<ArgumentException>(() => paragraph.AddHyperlink("a\nb", destination));
        Assert.Throws<ArgumentException>(() => paragraph.AddHyperlink("\0", destination));

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(paragraph.Hyperlinks);
        Assert.Empty(OoxmlValidation.ReadDocxExternalHyperlinkUris(produced));
    }

    [Fact]
    public void Hyperlink_RejectsInvalidEditsWithoutChangingTheLink()
    {
        Uri destination = new("https://example.com/valid");
        using WordDocument document = WordDocument.Create();
        Hyperlink link = document.AddParagraph(string.Empty).AddHyperlink("texto", destination);

        Assert.Throws<ArgumentNullException>(() => link.Text = null!);
        Assert.Throws<ArgumentException>(() => link.Text = "a\tb");
        Assert.Throws<ArgumentNullException>(() => link.Destination = null!);

        Assert.Equal("texto", link.Text);
        Assert.Equal(destination, link.Destination);
    }
}
