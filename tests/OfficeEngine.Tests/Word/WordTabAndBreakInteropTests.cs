using System;
using System.Linq;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Independent verification of the tab and break markup produced by OfficeEngine, using the Open XML
/// SDK and the raw package instead of the code under test.
/// </summary>
public sealed class WordTabAndBreakInteropTests
{
    [Fact]
    public void ProducedDocument_WithTabsAndBreaks_IsSchemaValid()
    {
        Assert.Empty(OoxmlValidation.ValidateDocx(CreateDocumentWithTabsAndBreaks()));
    }

    [Fact]
    public void ProducedDocument_TabStops_AreReadableByTheOpenXmlSdk()
    {
        W.ParagraphProperties? properties = OoxmlValidation.ReadDocxFirstParagraphProperties(CreateDocumentWithTabsAndBreaks());

        W.TabStop[] stops = properties!.GetFirstChild<W.Tabs>()!.Elements<W.TabStop>().ToArray();
        Assert.Equal(2, stops.Length);
        Assert.Equal(W.TabStopValues.Left, stops[0].Val!.Value);
        Assert.Equal(720, stops[0].Position!.Value);
        Assert.Null(stops[0].Leader);
        Assert.Equal(W.TabStopValues.Decimal, stops[1].Val!.Value);
        Assert.Equal(1440, stops[1].Position!.Value);
        Assert.Equal(W.TabStopLeaderCharValues.Dot, stops[1].Leader!.Value);
    }

    [Fact]
    public void ProducedDocument_WritesParagraphPropertiesInSchemaOrder()
    {
        string xml = TestPackage.ReadEntryText(CreateDocumentWithTabsAndBreaks(), "word/document.xml");

        int tabs = xml.IndexOf("<w:tabs", StringComparison.Ordinal);
        int spacing = xml.IndexOf("<w:spacing", StringComparison.Ordinal);
        int justification = xml.IndexOf("<w:jc", StringComparison.Ordinal);

        Assert.True(
            tabs >= 0 && spacing > tabs && justification > spacing,
            $"Unexpected paragraph property order: {xml}");
    }

    [Fact]
    public void ProducedDocument_RunContent_IsReadableByTheOpenXmlSdk()
    {
        W.Run run = OoxmlValidation.ReadDocxFirstRun(CreateDocumentWithTabsAndBreaks())!;

        Assert.Single(run.Elements<W.TabChar>());
        W.Break[] breaks = run.Elements<W.Break>().ToArray();
        Assert.Equal(2, breaks.Length);
        Assert.Equal(W.BreakValues.Page, breaks[0].Type!.Value);
        Assert.Null(breaks[1].Type);
    }

    [Fact]
    public void ProducedDocument_NeverWritesAnExplicitTextWrappingBreakType()
    {
        string xml = TestPackage.ReadEntryText(CreateDocumentWithTabsAndBreaks(), "word/document.xml");

        // A text-wrapping break is expressed by omitting @w:type, which is what Word does.
        Assert.Contains("w:type=\"page\"", xml);
        Assert.DoesNotContain("textWrapping", xml);
    }

    [Fact]
    public void AddingATabStop_KeepsUnmodelledParagraphMarkupInSchemaOrder()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithParagraphProperties("x", new W.ParagraphProperties(
            new W.KeepNext(),
            new W.Indentation { Left = "360" },
            new W.Justification { Val = W.JustificationValues.Center }));

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            document.Paragraphs[0].AddTabStop(new TabStop(18.0, TabStopAlignment.Right));
            produced = SaveBytes.ToBytes(document);
        }

        string xml = TestPackage.ReadEntryText(produced, "word/document.xml");
        int keepNext = xml.IndexOf("<w:keepNext", StringComparison.Ordinal);
        int tabs = xml.IndexOf("<w:tabs", StringComparison.Ordinal);
        int indent = xml.IndexOf("<w:ind", StringComparison.Ordinal);
        int justification = xml.IndexOf("<w:jc", StringComparison.Ordinal);

        Assert.True(
            keepNext >= 0 && tabs > keepNext && indent > tabs && justification > indent,
            $"Unexpected paragraph property order: {xml}");
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void ProducedDocument_WithPreExistingTabStops_RoundTripsThroughTheSdk()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTabStops(
            "x",
            new W.TabStop { Val = W.TabStopValues.Left, Position = 720 },
            new W.TabStop { Val = W.TabStopValues.Right, Position = 1440, Leader = W.TabStopLeaderCharValues.Dot });

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Paragraph paragraph = document.Paragraphs[0];
            Assert.Equal(2, paragraph.TabStops.Count);
            paragraph.AddTabStop(new TabStop(36.0, TabStopAlignment.Center));
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        W.ParagraphProperties? properties = OoxmlValidation.ReadDocxFirstParagraphProperties(produced);
        W.TabStop[] stops = properties!.GetFirstChild<W.Tabs>()!.Elements<W.TabStop>().ToArray();
        Assert.Equal(3, stops.Length);
        Assert.Equal(720, stops[0].Position!.Value);
        Assert.Equal(1440, stops[1].Position!.Value);
        Assert.Equal(720, stops[2].Position!.Value);
        Assert.Equal(W.TabStopValues.Center, stops[2].Val!.Value);
    }

    private static byte[] CreateDocumentWithTabsAndBreaks()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        paragraph.AddTabStop(new TabStop(36.0, TabStopAlignment.Left));
        paragraph.AddTabStop(new TabStop(72.0, TabStopAlignment.Decimal, TabStopLeader.Dot));
        paragraph.Alignment = ParagraphAlignment.Center;
        paragraph.SpacingBeforePoints = 6.0;

        Run run = paragraph.AddRun("A");
        run.AddTab().AddText("B").AddBreak(BreakType.Page).AddText("C");
        run.AddBreak(BreakType.TextWrapping);

        return SaveBytes.ToBytes(document);
    }
}
