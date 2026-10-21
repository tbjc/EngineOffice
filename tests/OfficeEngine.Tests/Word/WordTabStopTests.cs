using System;
using OfficeEngine;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Specification of paragraph tab stops (<c>w:pPr/w:tabs</c>) through the public API, including the
/// values the schema allows but this version does not model.
/// </summary>
public sealed class WordTabStopTests
{
    [Fact]
    public void Paragraph_TabStops_AreEmptyByDefault()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        Assert.Empty(paragraph.TabStops);
    }

    [Fact]
    public void Paragraph_AddTabStop_IsReadableBack()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        paragraph.AddTabStop(new TabStop(36.0, TabStopAlignment.Left));
        paragraph.AddTabStop(new TabStop(72.0, TabStopAlignment.Center, TabStopLeader.Dot));
        paragraph.AddTabStop(new TabStop(108.0, TabStopAlignment.Decimal, TabStopLeader.Underscore));

        Assert.Equal(
            new[]
            {
                new TabStop(36.0, TabStopAlignment.Left),
                new TabStop(72.0, TabStopAlignment.Center, TabStopLeader.Dot),
                new TabStop(108.0, TabStopAlignment.Decimal, TabStopLeader.Underscore),
            },
            paragraph.TabStops);
    }

    [Fact]
    public void Paragraph_AddTabStop_RoundsThePositionToTheNearestTwip()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        paragraph.AddTabStop(new TabStop(45.03, TabStopAlignment.Left));

        Assert.Equal(45.05, paragraph.TabStops[0].PositionPoints);
    }

    [Fact]
    public void Paragraph_AddTabStop_PreservesANegativePosition()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        paragraph.AddTabStop(new TabStop(-12.5, TabStopAlignment.Right));

        Assert.Equal(-12.5, paragraph.TabStops[0].PositionPoints);
    }

    [Fact]
    public void Paragraph_TabStops_RoundTripThroughSaveAndOpen()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            Paragraph paragraph = document.AddParagraph("tabbed");
            paragraph.AddTabStop(new TabStop(20.0, TabStopAlignment.Left));
            paragraph.AddTabStop(new TabStop(40.0, TabStopAlignment.Bar, TabStopLeader.Dot));
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Open(stream);
        Assert.Equal(
            new[] { new TabStop(20.0, TabStopAlignment.Left), new TabStop(40.0, TabStopAlignment.Bar, TabStopLeader.Dot) },
            reopened.Paragraphs[0].TabStops);
    }

    [Fact]
    public void Paragraph_ClearTabStops_RemovesEveryStop()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        paragraph.AddTabStop(new TabStop(20.0, TabStopAlignment.Left));

        paragraph.ClearTabStops();

        Assert.Empty(paragraph.TabStops);
    }

    [Fact]
    public void Paragraph_ClearTabStops_IsSafeWithoutParagraphProperties()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        paragraph.ClearTabStops();

        Assert.Empty(paragraph.TabStops);
    }

    [Fact]
    public void Paragraph_ClearingTabStops_RemovesTheEmptyProperties()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        paragraph.AddTabStop(new TabStop(20.0, TabStopAlignment.Left));

        paragraph.ClearTabStops();
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Null(OoxmlValidation.ReadDocxFirstParagraphProperties(produced));
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_AddTabStop_WritesEveryModelledAlignmentAndLeader()
    {
        TabStopAlignment[] alignments =
        {
            TabStopAlignment.Clear,
            TabStopAlignment.Left,
            TabStopAlignment.Center,
            TabStopAlignment.Right,
            TabStopAlignment.Decimal,
            TabStopAlignment.Bar,
        };

        TabStopLeader[] leaders =
        {
            TabStopLeader.None,
            TabStopLeader.Dot,
            TabStopLeader.Hyphen,
            TabStopLeader.Underscore,
            TabStopLeader.Heavy,
            TabStopLeader.MiddleDot,
        };

        List<TabStop> expected = new();
        double position = 0.0;
        foreach (TabStopAlignment alignment in alignments)
        {
            foreach (TabStopLeader leader in leaders)
            {
                position += 1.0;
                expected.Add(new TabStop(position, alignment, leader));
            }
        }

        byte[] produced;
        using (WordDocument document = WordDocument.Create())
        {
            Paragraph paragraph = document.AddParagraph("tabbed");
            foreach (TabStop stop in expected)
            {
                paragraph.AddTabStop(stop);
            }

            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Equal(expected, reopened.Paragraphs[0].TabStops);
    }

    [Fact]
    public void Paragraph_AddTabStop_KeepsOtherFormatting()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        paragraph.Alignment = ParagraphAlignment.Center;

        paragraph.AddTabStop(new TabStop(20.0, TabStopAlignment.Left));

        Assert.Equal(ParagraphAlignment.Center, paragraph.Alignment);
        Assert.Single(paragraph.TabStops);
    }

    [Fact]
    public void Paragraph_AddTabStop_WithoutLeader_OmitsTheLeaderAttribute()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        paragraph.AddTabStop(new TabStop(20.0, TabStopAlignment.Left));

        W.ParagraphProperties? properties = OoxmlValidation.ReadDocxFirstParagraphProperties(SaveBytes.ToBytes(document));

        W.TabStop stop = Assert.Single(properties!.GetFirstChild<W.Tabs>()!.Elements<W.TabStop>());
        Assert.Equal(400, stop.Position!.Value);
        Assert.Null(stop.Leader);
    }

    [Fact]
    public void Paragraph_AddTabStopWithTooLargePosition_ThrowsAndLeavesNoMarkup()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        Assert.Throws<ArgumentOutOfRangeException>(
            () => paragraph.AddTabStop(new TabStop(double.MaxValue, TabStopAlignment.Left)));

        Assert.Empty(paragraph.TabStops);
    }

    [Fact]
    public void Paragraph_StoredUnmodelledTabStop_ThrowsAndIsPreserved()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredTabStop("x", "num", "720", null);
        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Paragraph paragraph = document.Paragraphs[0];
            Assert.Throws<UnsupportedOfficeFeatureException>(() => paragraph.TabStops);

            // The unmodelled stop must not be silently rewritten or dropped.
            paragraph.SpacingBeforePoints = 6.0;
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Contains("w:val=\"num\"", TestPackage.ReadEntryText(produced, "word/document.xml"));
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_TabStops_MissingAlignment_ThrowsInvalidOfficeFile()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredTabStop("x", null, "720", null);
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].TabStops);
    }

    [Fact]
    public void Paragraph_TabStops_MissingPosition_ThrowsInvalidOfficeFile()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredTabStop("x", "left", null, null);
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].TabStops);
    }

    [Fact]
    public void Paragraph_TabStops_MalformedPosition_ThrowsInvalidOfficeFile()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredTabStop("x", "left", "not-a-number", null);
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].TabStops);
    }

    [Fact]
    public void Paragraph_TabStops_MalformedAlignment_ThrowsInvalidOfficeFile()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredTabStop("x", "bogus", "720", null);
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].TabStops);
    }

    [Fact]
    public void Paragraph_TabStops_MalformedLeader_ThrowsInvalidOfficeFile()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredTabStop("x", "left", "720", "bogus");
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].TabStops);
    }

    [Fact]
    public void Paragraph_TabStops_ReadEveryModelledAlignmentAndLeader()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTabStops(
            "x",
            new W.TabStop { Val = W.TabStopValues.Clear, Position = 100 },
            new W.TabStop { Val = W.TabStopValues.Left, Position = 200, Leader = W.TabStopLeaderCharValues.Dot },
            new W.TabStop { Val = W.TabStopValues.Center, Position = 300, Leader = W.TabStopLeaderCharValues.Hyphen },
            new W.TabStop { Val = W.TabStopValues.Right, Position = 400, Leader = W.TabStopLeaderCharValues.Underscore },
            new W.TabStop { Val = W.TabStopValues.Decimal, Position = 500, Leader = W.TabStopLeaderCharValues.Heavy },
            new W.TabStop { Val = W.TabStopValues.Bar, Position = 600, Leader = W.TabStopLeaderCharValues.MiddleDot });

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Equal(
            new[]
            {
                new TabStop(5.0, TabStopAlignment.Clear),
                new TabStop(10.0, TabStopAlignment.Left, TabStopLeader.Dot),
                new TabStop(15.0, TabStopAlignment.Center, TabStopLeader.Hyphen),
                new TabStop(20.0, TabStopAlignment.Right, TabStopLeader.Underscore),
                new TabStop(25.0, TabStopAlignment.Decimal, TabStopLeader.Heavy),
                new TabStop(30.0, TabStopAlignment.Bar, TabStopLeader.MiddleDot),
            },
            document.Paragraphs[0].TabStops);
    }

    [Fact]
    public void TabStop_Constructor_RejectsNonFinitePosition()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new TabStop(double.NaN, TabStopAlignment.Left));
        Assert.Throws<ArgumentOutOfRangeException>(() => new TabStop(double.PositiveInfinity, TabStopAlignment.Left));
        Assert.Throws<ArgumentOutOfRangeException>(() => new TabStop(double.NegativeInfinity, TabStopAlignment.Left));
    }

    [Fact]
    public void Paragraph_AddTabStop_RejectsAnUndefinedAlignment()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.AddTabStop(new TabStop(1.0, (TabStopAlignment)99)));

        Assert.Empty(paragraph.TabStops);
    }

    [Fact]
    public void Paragraph_AddTabStop_RejectsAnUndefinedLeader()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        Assert.Throws<ArgumentOutOfRangeException>(
            () => paragraph.AddTabStop(new TabStop(1.0, TabStopAlignment.Left, (TabStopLeader)99)));

        Assert.Empty(paragraph.TabStops);
    }

    [Fact]
    public void TabStop_EqualityAndFormatting()
    {
        TabStop first = new(10.0, TabStopAlignment.Right, TabStopLeader.Dot);
        TabStop same = new(10.0, TabStopAlignment.Right, TabStopLeader.Dot);

        Assert.True(first == same);
        Assert.False(first != same);
        Assert.Equal(first, same);
        Assert.Equal(first.GetHashCode(), same.GetHashCode());
        Assert.NotEqual(first, new TabStop(11.0, TabStopAlignment.Right, TabStopLeader.Dot));
        Assert.NotEqual(first, new TabStop(10.0, TabStopAlignment.Left, TabStopLeader.Dot));
        Assert.NotEqual(first, new TabStop(10.0, TabStopAlignment.Right));
        Assert.False(first.Equals(null));
        Assert.Equal("Right at 10 pt (Dot)", first.ToString());
    }
}
