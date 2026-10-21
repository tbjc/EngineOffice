using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.11: search and replace. The scope is the body, the search never crosses a paragraph
/// and the replaced text keeps the formatting of the first run of the match.
/// </summary>
public sealed class WordTextSearchTests
{
    [Fact]
    public void FindText_FindsEveryOccurrenceInTheBody()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph first = document.AddParagraph("La versión borrador sigue");
        Paragraph second = document.AddParagraph("Otra versión borrador aquí");

        IReadOnlyList<WordTextMatch> matches = document.FindText("borrador");

        Assert.Equal(2, matches.Count);
        Assert.Equal(first.Text, matches[0].Paragraph.Text);
        Assert.Equal(11, matches[0].StartIndex);
        Assert.Equal("borrador", matches[0].Text);
        Assert.Equal(second.Text, matches[1].Paragraph.Text);
        Assert.Equal(13, matches[1].StartIndex);
    }

    [Fact]
    public void FindText_WithTextSplitAcrossRuns_FindsIt()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("bo");
        paragraph.AddRun("rrador");

        IReadOnlyList<WordTextMatch> matches = document.FindText("borrador");

        Assert.Single(matches);
        Assert.Equal(0, matches[0].StartIndex);
        Assert.Equal("borrador", matches[0].Text);
    }

    [Fact]
    public void FindText_ComparesCaseOnlyWhenAsked()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Borrador y borrador");

        Assert.Equal(2, document.FindText("borrador").Count);
        Assert.Single(document.FindText("borrador", matchCase: true));
    }

    [Fact]
    public void FindText_DoesNotCrossParagraphs()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("borra");
        document.AddParagraph("dor");

        Assert.Empty(document.FindText("borrador"));
    }

    [Fact]
    public void FindText_DoesNotSearchHeadersOrFooters()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("cuerpo");
        HeaderFooter header = document.Sections[0].AddHeader(HeaderFooterType.Primary);
        header.Paragraphs[0].Text = "borrador en el encabezado";

        Assert.Empty(document.FindText("borrador"));
    }

    [Fact]
    public void ReplaceText_ReplacesEveryOccurrenceAndKeepsTheFormatOfTheFirstRun()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("bo");
        Run second = paragraph.AddRun("rrador final");
        second.Bold = true;

        int replaced = document.ReplaceText("borrador", "final");

        Assert.Equal(1, replaced);
        Assert.Equal("final final", paragraph.Text);
        Assert.Equal(2, paragraph.Runs.Count);
        Assert.Null(paragraph.Runs[0].Bold);
        Assert.True(paragraph.Runs[1].Bold);
    }

    [Fact]
    public void ReplaceText_KeepsTheTextAroundTheMatch()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("antes ");
        paragraph.AddRun("borrador");
        paragraph.AddRun(" después");

        document.ReplaceText("borrador", "nuevo");

        Assert.Equal("antes nuevo después", paragraph.Text);
    }

    [Fact]
    public void ReplaceText_WithAnEmptyNewText_RemovesTheText()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("borra");
        paragraph.AddRun("borrador");

        document.ReplaceText("borrador", string.Empty);

        Assert.Equal("borra", paragraph.Text);
    }

    [Fact]
    public void ReplaceText_ReachesTheTextOfTables()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "borrador aquí";

        int replaced = document.ReplaceText("borrador", "final");

        Assert.Equal(1, replaced);
        Assert.Equal("final aquí", table.Rows[0].Cells[0].Paragraphs[0].Text);
    }

    [Fact]
    public void ReplaceText_RoundTripsThroughSaveAndOpen()
    {
        byte[] produced;
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("un borrador");
            document.ReplaceText("borrador", "final");
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Equal("un final", reopened.Paragraphs[0].Text);
    }

    [Fact]
    public void ReplaceText_KeepsTheRestOfTheDocument()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("borrador");
        paragraph.Alignment = ParagraphAlignment.Center;
        paragraph.AddRun(" y más");

        document.ReplaceText("borrador", "final");

        Assert.Equal(ParagraphAlignment.Center, paragraph.Alignment);
        Assert.Equal("final y más", paragraph.Text);
    }

    [Fact]
    public void ReplaceText_WithInvalidValues_Throws()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("borrador");

        Assert.Throws<ArgumentException>(() => document.ReplaceText(string.Empty, "x"));
        Assert.Throws<ArgumentNullException>(() => document.ReplaceText("borrador", null!));
        Assert.Throws<ArgumentException>(() => document.FindText(string.Empty));
        Assert.Equal("borrador", document.Paragraphs[0].Text);
    }

    [Fact]
    public void ReplaceText_KeepsTheWordOfTheDocument()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("borrador");

        document.ReplaceText("borrador", "borrador");

        Assert.Single(document.FindText("borrador"));
    }

    [Fact]
    public void ReplaceText_ProducesAValidPackage()
    {
        byte[] produced;
        using (WordDocument document = WordDocument.Create())
        {
            Paragraph paragraph = document.AddParagraph("bo");
            paragraph.AddRun("rrador");
            document.ReplaceText("borrador", "final");
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(
            "final",
            OoxmlValidation.ReadDocxText(produced).Split('\n')[0].Trim());
    }

    [Fact]
    public void Paragraph_FindText_IsScopedToThatParagraph()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph first = document.AddParagraph("uno");
        Paragraph second = document.AddParagraph("dos");

        Assert.Single(first.FindText("uno"));
        Assert.Empty(second.FindText("uno"));
        Assert.Equal(1, first.ReplaceText("uno", "1"));
        Assert.Equal(0, second.ReplaceText("uno", "1"));
    }

    [Fact]
    public void ReplacedText_IsNotLeftInTheOldRuns()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("pre ");
        Run run = paragraph.AddRun("borrador");
        run.Italic = true;
        paragraph.AddRun(" post");

        document.ReplaceText("borrador", "x");

        Assert.Equal("pre x post", paragraph.Text);
        Assert.DoesNotContain(paragraph.Runs, candidate => candidate.Text.Contains("borrador", StringComparison.Ordinal));
    }
}
