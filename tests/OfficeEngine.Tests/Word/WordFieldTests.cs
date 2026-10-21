using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.12: fields and page numbering. The fields are written, never calculated; the cached
/// result is what a reader shows until Word recomputes it.
/// </summary>
public sealed class WordFieldTests
{
    [Fact]
    public void AddPageNumber_WritesTheFieldInTheFooter()
    {
        byte[] produced;
        using (WordDocument document = WordDocument.Create())
        {
            HeaderFooter footer = document.Sections[0].AddFooter(HeaderFooterType.Primary);
            Paragraph paragraph = footer.AddPageNumber("1");

            Assert.Equal(5, paragraph.Runs.Count);
            Assert.Equal("1", paragraph.Text);

            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        (IReadOnlyList<string> Instructions, IReadOnlyList<string> FieldCharTypes) =
            OoxmlValidation.ReadDocxFooterFields(produced);
        Assert.Equal([" PAGE "], Instructions);
        Assert.Equal(["begin", "separate", "end"], FieldCharTypes);
    }

    [Fact]
    public void AddField_WithoutAResult_WritesOnlyTheFieldCharacters()
    {
        byte[] produced;
        using (WordDocument document = WordDocument.Create())
        {
            Paragraph paragraph = document.AddParagraph("Página ");
            paragraph.AddField("NUMPAGES");

            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        // The library reads the visible text: the instruction of the field is not part of it.
        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Equal("Página ", reopened.Paragraphs[0].Text);
        Assert.Contains("NUMPAGES", OoxmlValidation.ReadDocxText(produced), StringComparison.Ordinal);
    }

    [Fact]
    public void AddField_KeepsTheTextAroundIt()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("Página ");
        paragraph.AddField("PAGE", "3");
        paragraph.AddRun(" de 10");

        Assert.Equal("Página 3 de 10", paragraph.Text);
    }

    [Fact]
    public void AddField_RoundTripsThroughSaveAndOpen()
    {
        byte[] produced;
        using (WordDocument document = WordDocument.Create())
        {
            HeaderFooter footer = document.Sections[0].AddFooter(HeaderFooterType.Primary);
            footer.AddPageNumber("7");
            produced = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        HeaderFooter? reopenedFooter = reopened.Sections[0].Footer(HeaderFooterType.Primary);
        Assert.NotNull(reopenedFooter);
        Assert.Equal("7", reopenedFooter!.Text);
    }

    [Fact]
    public void AddField_IsNotCalculated()
    {
        byte[] produced;
        using (WordDocument document = WordDocument.Create())
        {
            HeaderFooter created = document.Sections[0].AddFooter(HeaderFooterType.Primary);
            created.AddPageNumber("1");
            produced = SaveBytes.ToBytes(document);
        }

        // The stored result is the one that was written; nobody recalculated it.
        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        HeaderFooter? reopenedFooter = reopened.Sections[0].Footer(HeaderFooterType.Primary);
        Assert.NotNull(reopenedFooter);
        Assert.Equal("1", reopenedFooter!.Text);
        Assert.Contains(
            " PAGE ",
            OoxmlValidation.ReadDocxFooterFields(produced).Instructions[0],
            StringComparison.Ordinal);
    }

    [Fact]
    public void AddField_WithAnInvalidInstruction_Throws()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("texto");

        Assert.Throws<ArgumentException>(() => paragraph.AddField(string.Empty));
        Assert.Throws<ArgumentException>(() => paragraph.AddField("PAGE\tNUMPAGES"));
        Assert.Equal("texto", paragraph.Text);
    }

    [Fact]
    public void AddField_WithAnInvalidResult_Throws()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("texto");

        Assert.Throws<ArgumentException>(() => paragraph.AddField("PAGE", "con\nsalto"));
    }

    [Fact]
    public void AddPageNumber_KeepsTheExistingParagraphOfTheFooter()
    {
        using WordDocument document = WordDocument.Create();
        HeaderFooter footer = document.Sections[0].AddFooter(HeaderFooterType.Primary);
        footer.Paragraphs[0].Text = "Informe";

        footer.AddPageNumber("1");

        Assert.Equal(2, footer.Paragraphs.Count);
        Assert.Equal("Informe\n1", footer.Text);
    }
}
