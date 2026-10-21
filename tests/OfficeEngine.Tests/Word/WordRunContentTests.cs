using System;
using System.Linq;
using OfficeEngine;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Specification of the modelled content of a run (text, tabs and breaks) and of the rejection of
/// characters that plain text cannot carry.
/// </summary>
public sealed class WordRunContentTests
{
    [Fact]
    public void Run_Content_ReportsText()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph("hola").Runs[0];

        RunText text = Assert.IsType<RunText>(Assert.Single(run.Content));
        Assert.Equal("hola", text.Value);
        Assert.Equal("hola", text.ToString());
    }

    [Fact]
    public void Run_AddText_AppendsTextAndReturnsTheSameRun()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun(string.Empty);

        Run returned = run.AddText("a").AddText("b");

        Assert.Same(run, returned);
        Assert.Equal("ab", run.Text);
        Assert.Equal(2, run.Content.Count);
    }

    [Fact]
    public void Run_AddText_IgnoresAnEmptyString()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun(string.Empty);

        run.AddText(string.Empty);

        Assert.Empty(run.Content);
    }

    [Fact]
    public void Run_AddTab_AppendsATabCharacter()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun(string.Empty);

        Run returned = run.AddTab();

        Assert.Same(run, returned);
        RunTab tab = Assert.IsType<RunTab>(Assert.Single(run.Content));
        Assert.Equal("\t", tab.ToString());
    }

    [Fact]
    public void Run_AddBreak_AppendsEachModelledBreakType()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun(string.Empty);

        run.AddBreak(BreakType.TextWrapping).AddBreak(BreakType.Page).AddBreak(BreakType.Column);

        Assert.Collection(
            run.Content,
            item => Assert.Equal(BreakType.TextWrapping, Assert.IsType<RunBreak>(item).Type),
            item => Assert.Equal(BreakType.Page, Assert.IsType<RunBreak>(item).Type),
            item => Assert.Equal(BreakType.Column, Assert.IsType<RunBreak>(item).Type));
    }

    [Fact]
    public void Run_AddBreak_RejectsAnUndefinedType()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun(string.Empty);

        Assert.Throws<ArgumentOutOfRangeException>(() => run.AddBreak((BreakType)99));
        Assert.Empty(run.Content);
    }

    [Fact]
    public void Run_Content_ReportsContentInDocumentOrder()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun(string.Empty);
        run.AddText("A").AddTab().AddText("B").AddBreak(BreakType.Page).AddText("C");

        Assert.Collection(
            run.Content,
            item => Assert.Equal("A", Assert.IsType<RunText>(item).Value),
            item => Assert.IsType<RunTab>(item),
            item => Assert.Equal("B", Assert.IsType<RunText>(item).Value),
            item => Assert.Equal(BreakType.Page, Assert.IsType<RunBreak>(item).Type),
            item => Assert.Equal("C", Assert.IsType<RunText>(item).Value));

        // Text keeps its documented meaning: it concatenates w:t only.
        Assert.Equal("ABC", run.Text);
        Assert.Equal(BreakType.Page.ToString(), run.Content[3].ToString());
    }

    [Fact]
    public void Run_TextSetter_ReplacesTabsAndBreaks()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun(string.Empty);
        run.AddTab().AddText("x").AddBreak(BreakType.Page);

        run.Text = "solo";

        RunText text = Assert.IsType<RunText>(Assert.Single(run.Content));
        Assert.Equal("solo", text.Value);
    }

    [Fact]
    public void Run_Content_RoundTripsThroughSaveAndOpen()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            Run run = document.AddParagraph(string.Empty).AddRun("A");
            run.AddTab().AddText("B").AddBreak(BreakType.Page).AddText("C").AddBreak(BreakType.Column);
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Open(stream);
        Run reloaded = reopened.Paragraphs[0].Runs[0];

        Assert.Equal("ABC", reloaded.Text);
        Assert.Collection(
            reloaded.Content,
            item => Assert.Equal("A", Assert.IsType<RunText>(item).Value),
            item => Assert.IsType<RunTab>(item),
            item => Assert.Equal("B", Assert.IsType<RunText>(item).Value),
            item => Assert.Equal(BreakType.Page, Assert.IsType<RunBreak>(item).Type),
            item => Assert.Equal("C", Assert.IsType<RunText>(item).Value),
            item => Assert.Equal(BreakType.Column, Assert.IsType<RunBreak>(item).Type));
    }

    [Fact]
    public void Run_AddText_PreservesSurroundingWhitespace()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph(string.Empty).AddRun(string.Empty).AddText("  spaced  ");
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Open(stream);
        Assert.Equal("  spaced  ", reopened.Paragraphs[0].Runs[0].Text);
    }

    [Fact]
    public void Run_Content_MapsACarriageReturnToATextWrappingBreak()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithMixedRunContent("head");
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Run run = document.Paragraphs[0].Runs[0];

        // w:cr is a line break, and w:lastRenderedPageBreak is a rendering hint that is not content.
        Assert.Collection(
            run.Content,
            item => Assert.Equal("head", Assert.IsType<RunText>(item).Value),
            item => Assert.Equal(BreakType.TextWrapping, Assert.IsType<RunBreak>(item).Type),
            item => Assert.IsType<RunTab>(item),
            item => Assert.Equal("tail", Assert.IsType<RunText>(item).Value));
    }

    [Fact]
    public void Run_Content_ReadsEachModelledStoredBreakType()
    {
        Assert.Equal(BreakType.Column, ReadStoredBreakType(OoxmlValidation.CreateDocxWithStoredBreak("x", "column")));
        Assert.Equal(BreakType.Page, ReadStoredBreakType(OoxmlValidation.CreateDocxWithStoredBreak("x", "page")));
        Assert.Equal(BreakType.TextWrapping, ReadStoredBreakType(OoxmlValidation.CreateDocxWithStoredBreak("x", "textWrapping")));

        // An omitted @w:type means a text-wrapping break.
        Assert.Equal(BreakType.TextWrapping, ReadStoredBreakType(OoxmlValidation.CreateDocxWithStoredBreak("x", null)));
    }

    [Fact]
    public void Run_Content_MalformedStoredBreakType_ThrowsInvalidOfficeFile()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredBreak("x", "bogus");
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].Runs[0].Content);
    }

    [Fact]
    public void Run_Editing_LeavesUnmodelledRunContentUntouched()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithMixedRunContent("keep");
        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            document.Paragraphs[0].Runs[0].AddTab();
            produced = SaveBytes.ToBytes(document);
        }

        string xml = TestPackage.ReadEntryText(produced, "word/document.xml");
        Assert.Contains("<w:cr", xml);
        Assert.Contains("lastRenderedPageBreak", xml);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Theory]
    [InlineData("a\tb")]
    [InlineData("a\nb")]
    [InlineData("a\rb")]
    public void Run_TextSetter_RejectsTabsAndLineBreaks(string value)
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph("keep").Runs[0];

        Assert.Throws<ArgumentException>(() => run.Text = value);

        Assert.Equal("keep", run.Text);
    }

    [Theory]
    [InlineData("a\tb")]
    [InlineData("a\nb")]
    public void Run_AddText_RejectsTabsAndLineBreaks(string value)
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun(string.Empty);

        Assert.Throws<ArgumentException>(() => run.AddText(value));
        Assert.Empty(run.Content);
    }

    [Fact]
    public void Paragraph_AddRun_RejectsTabsAndLineBreaks()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("x");

        Assert.Throws<ArgumentException>(() => paragraph.AddRun("a\nb"));
        Assert.Single(paragraph.Runs);
    }

    [Fact]
    public void Paragraph_AppendText_RejectsTabsAndLineBreaks()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("x");

        Assert.Throws<ArgumentException>(() => paragraph.AppendText("a\tb"));
        Assert.Equal("x", paragraph.Text);
    }

    [Fact]
    public void Paragraph_TextSetter_RejectsTabsAndLineBreaks()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("x");

        Assert.Throws<ArgumentException>(() => paragraph.Text = "a\tb");
        Assert.Equal("x", paragraph.Text);
    }

    [Fact]
    public void WordDocument_AddParagraph_RejectsTabsAndLineBreaks()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Throws<ArgumentException>(() => document.AddParagraph("a\nb"));
        Assert.Empty(document.Paragraphs);
    }

    [Fact]
    public void Run_AddText_RejectsCharactersOfficeOpenXmlCannotStore()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun(string.Empty);

        Assert.Throws<ArgumentException>(() => run.AddText("a\u0000b"));
        Assert.Throws<ArgumentException>(() => run.AddText("a\u0007b"));
        Assert.Empty(run.Content);
    }

    [Fact]
    public void Run_AddText_RejectsALoneSurrogate()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun(string.Empty);

        // A high surrogate followed by a normal character, a trailing high surrogate, and a low
        // surrogate on its own are all unrepresentable in XML.
        Assert.Throws<ArgumentException>(() => run.AddText("a\uD83Db"));
        Assert.Throws<ArgumentException>(() => run.AddText("a\uD83D"));
        Assert.Throws<ArgumentException>(() => run.AddText("a\uDE00b"));
        Assert.Empty(run.Content);
    }

    [Fact]
    public void Run_AddText_AcceptsAValidSurrogatePair()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun(string.Empty);

        run.AddText("a\uD83D\uDE00b");

        Assert.Equal("a\uD83D\uDE00b", run.Text);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void Run_AddText_NullThrows()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun(string.Empty);

        Assert.Throws<ArgumentNullException>(() => run.AddText(null!));
    }

    private static BreakType ReadStoredBreakType(byte[] fixture)
    {
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        return document.Paragraphs[0].Runs[0].Content.OfType<RunBreak>().Single().Type;
    }
}
