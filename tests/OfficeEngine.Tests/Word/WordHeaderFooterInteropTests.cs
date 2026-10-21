using System.Linq;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.7: interoperability of headers and footers, verified with the Open XML SDK instead of the
/// code under test — produced packages are read back and validated independently, and fixtures built with
/// the SDK are read by OfficeEngine.
/// </summary>
public sealed class WordHeaderFooterInteropTests
{
    [Fact]
    public void ProducedDocument_WithHeaderAndFooter_IsSchemaValidAndLinked()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];
        section.AddHeader(HeaderFooterType.Primary).Paragraphs[0].Text = "Cabecera";
        section.AddFooter(HeaderFooterType.Primary).Paragraphs[0].Text = "Pie";
        section.DifferentFirstPage = true;
        section.AddHeader(HeaderFooterType.FirstPage).Paragraphs[0].Text = "Primera";

        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        W.SectionProperties stored = OoxmlValidation.ReadDocxFinalSectionProperties(produced)!;
        var headerReferences = stored.Elements<W.HeaderReference>().ToArray();
        Assert.Equal(2, headerReferences.Length);
        Assert.Contains(headerReferences, reference => reference.Type!.Value == W.HeaderFooterValues.Default);
        Assert.Contains(headerReferences, reference => reference.Type!.Value == W.HeaderFooterValues.First);

        W.FooterReference footerReference = Assert.Single(stored.Elements<W.FooterReference>());
        Assert.Equal(W.HeaderFooterValues.Default, footerReference.Type!.Value);

        Assert.Equal(2, OoxmlValidation.CountDocxHeaderParts(produced));
        Assert.Equal(1, OoxmlValidation.CountDocxFooterParts(produced));
        Assert.Equal("Pie", Assert.Single(OoxmlValidation.ReadDocxFooterTexts(produced)));
    }

    [Fact]
    public void ProducedDocument_WithEveryHeaderAndFooterType_IsSchemaValidWithReferencesInOrder()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];
        foreach (HeaderFooterType type in System.Enum.GetValues<HeaderFooterType>())
        {
            section.AddHeader(type).AddParagraph($"cabecera {type}");
            section.AddFooter(type).AddParagraph($"pie {type}");
        }

        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(3, OoxmlValidation.CountDocxHeaderParts(produced));
        Assert.Equal(3, OoxmlValidation.CountDocxFooterParts(produced));

        W.SectionProperties stored = OoxmlValidation.ReadDocxFinalSectionProperties(produced)!;
        Assert.Equal(
            [
                "headerReference", "headerReference", "headerReference",
                "footerReference", "footerReference", "footerReference",
                "pgSz", "pgMar",
            ],
            stored.ChildElements.Select(child => child.LocalName));
    }

    [Fact]
    public void OpenedDocument_WithTwoSectionsSharingAHeader_EditsTheSharedPart()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTwoSectionsSharingAHeader("uno", "dos", "Compartido");

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Assert.Equal("Compartido", document.Sections[0].Header(HeaderFooterType.Primary)!.Text);
            Assert.Equal("Compartido", document.Sections[1].Header(HeaderFooterType.Primary)!.Text);

            document.Sections[0].Header(HeaderFooterType.Primary)!.Paragraphs[0].Text = "Cambiado";
            Assert.Equal("Cambiado", document.Sections[1].Header(HeaderFooterType.Primary)!.Text);
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(1, OoxmlValidation.CountDocxHeaderParts(produced));
        Assert.Equal("Cambiado", Assert.Single(OoxmlValidation.ReadDocxHeaderTexts(produced)));
    }

    [Fact]
    public void OpenedDocument_WithUnmodelledHeaderContent_PreservesItWhenTheTextIsEdited()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithUnmodelledHeaderContent("cuerpo", "Cabecera");

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Assert.Equal("Cabecera", document.Sections[0].Header(HeaderFooterType.Primary)!.Text);
            document.Sections[0].Header(HeaderFooterType.Primary)!.Paragraphs[0].Text = "Nueva";
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal("Nueva", Assert.Single(OoxmlValidation.ReadDocxHeaderTexts(produced)));
        Assert.True(OoxmlValidation.ReadDocxHeaderHasContentControl(produced));
    }

    [Fact]
    public void OpenedDocument_WithAnOrphanHeaderPart_PreservesIt()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithOrphanHeader("cuerpo", "Huérfano");

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            document.AddParagraph("otro");
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal("Huérfano", Assert.Single(OoxmlValidation.ReadDocxHeaderTexts(produced)));
    }

    [Fact]
    public void OpenedDocument_RemovingOneReference_KeepsTheSharedPartForTheOtherSection()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTwoSectionsSharingAHeader("uno", "dos", "Compartido");

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Assert.True(document.Sections[0].RemoveHeader(HeaderFooterType.Primary));
            Assert.Null(document.Sections[0].Header(HeaderFooterType.Primary));
            Assert.Equal("Compartido", document.Sections[1].Header(HeaderFooterType.Primary)!.Text);
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(1, OoxmlValidation.CountDocxHeaderParts(produced));

        var sections = OoxmlValidation.ReadDocxSections(produced);
        Assert.Empty(sections[0].Elements<W.HeaderReference>());
        Assert.Single(sections[1].Elements<W.HeaderReference>());
    }

    [Fact]
    public void OpenedDocument_WithAHeaderReferenceToAFooterPart_Throws()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithHeaderReferenceToFooterPart("cuerpo");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Sections[0].Header(HeaderFooterType.Primary));
    }

    [Fact]
    public void ProducedDocument_AfterRemovingEveryReference_IsSchemaValid()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];
        section.AddHeader(HeaderFooterType.Primary);
        section.AddFooter(HeaderFooterType.EvenPage);
        section.RemoveHeader(HeaderFooterType.Primary);
        section.RemoveFooter(HeaderFooterType.EvenPage);

        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        W.SectionProperties stored = OoxmlValidation.ReadDocxFinalSectionProperties(produced)!;
        Assert.Empty(stored.Elements<W.HeaderReference>());
        Assert.Empty(stored.Elements<W.FooterReference>());
    }
}
