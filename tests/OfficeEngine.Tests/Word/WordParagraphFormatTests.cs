using System;
using OfficeEngine;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Word;

public sealed class WordParagraphFormatTests
{
    [Fact]
    public void Paragraph_DefaultAlignment_IsNull()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        Assert.Null(paragraph.Alignment);
    }

    [Fact]
    public void Paragraph_Alignment_CanBeSetAndRetrieved()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        paragraph.Alignment = ParagraphAlignment.Center;
        Assert.Equal(ParagraphAlignment.Center, paragraph.Alignment);

        paragraph.Alignment = null;
        Assert.Null(paragraph.Alignment);
    }

    [Fact]
    public void Paragraph_Alignment_RoundTripsThroughSaveAndOpen()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("alineado").Alignment = ParagraphAlignment.Right;
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Open(stream);
        Assert.Equal(ParagraphAlignment.Right, reopened.Paragraphs[0].Alignment);
    }

    [Fact]
    public void Paragraph_Alignment_AllValuesRoundTrip()
    {
        ParagraphAlignment[] values = new[] { ParagraphAlignment.Left, ParagraphAlignment.Center, ParagraphAlignment.Right, ParagraphAlignment.Justify, ParagraphAlignment.Both };
        foreach (ParagraphAlignment alignment in values)
        {
            using MemoryStream stream = new();
            using (WordDocument document = WordDocument.Create())
            {
                document.AddParagraph("x").Alignment = alignment;
                document.Save(stream);
            }

            stream.Position = 0;
            using WordDocument reopened = WordDocument.Open(stream);
            Assert.Equal(alignment, reopened.Paragraphs[0].Alignment);
        }
    }

    [Fact]
    public void Paragraph_Indents_CanBeSetAndCleared()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("indented");

        paragraph.FirstLineIndentPoints = 12.5;
        paragraph.LeftIndentPoints = 10.0;
        paragraph.RightIndentPoints = 5.0;
        paragraph.HangingIndentPoints = 6.0;

        Assert.Equal(12.5, paragraph.FirstLineIndentPoints);
        Assert.Equal(10.0, paragraph.LeftIndentPoints);
        Assert.Equal(5.0, paragraph.RightIndentPoints);
        Assert.Equal(6.0, paragraph.HangingIndentPoints);

        paragraph.FirstLineIndentPoints = null;
        paragraph.LeftIndentPoints = null;
        paragraph.RightIndentPoints = null;
        paragraph.HangingIndentPoints = null;

        Assert.Null(paragraph.FirstLineIndentPoints);
        Assert.Null(paragraph.LeftIndentPoints);
        Assert.Null(paragraph.RightIndentPoints);
        Assert.Null(paragraph.HangingIndentPoints);
    }

    [Fact]
    public void Paragraph_Indents_RoundTrip()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            Paragraph p = document.AddParagraph("ind");
            p.FirstLineIndentPoints = 12.0;
            p.LeftIndentPoints = 24.0;
            p.HangingIndentPoints = 12.0;
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Open(stream);
        Paragraph rp = reopened.Paragraphs[0];
        Assert.Equal(12.0, rp.FirstLineIndentPoints);
        Assert.Equal(24.0, rp.LeftIndentPoints);
        Assert.Equal(12.0, rp.HangingIndentPoints);
        Assert.Null(rp.RightIndentPoints);
    }

    [Fact]
    public void Paragraph_Spacing_CanBeSetAndCleared()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("spaced");

        paragraph.SpacingBeforePoints = 6.0;
        paragraph.SpacingAfterPoints = 12.0;
        paragraph.LineSpacingPoints = 18.0;

        Assert.Equal(6.0, paragraph.SpacingBeforePoints);
        Assert.Equal(12.0, paragraph.SpacingAfterPoints);
        Assert.Equal(18.0, paragraph.LineSpacingPoints);

        paragraph.SpacingBeforePoints = null;
        paragraph.SpacingAfterPoints = null;
        paragraph.LineSpacingPoints = null;

        Assert.Null(paragraph.SpacingBeforePoints);
        Assert.Null(paragraph.SpacingAfterPoints);
        Assert.Null(paragraph.LineSpacingPoints);
    }

    [Fact]
    public void Paragraph_Spacing_RoundTrip()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            Paragraph p = document.AddParagraph("sp");
            p.SpacingBeforePoints = 10.0;
            p.SpacingAfterPoints = 15.0;
            p.LineSpacingPoints = 12.0;
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Open(stream);
        Paragraph rp = reopened.Paragraphs[0];
        Assert.Equal(10.0, rp.SpacingBeforePoints);
        Assert.Equal(15.0, rp.SpacingAfterPoints);
        Assert.Equal(12.0, rp.LineSpacingPoints);
    }

    [Fact]
    public void Paragraph_RejectsInvalidNumericValues()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("x");

        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.FirstLineIndentPoints = -double.Epsilon);
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.LeftIndentPoints = double.NaN);
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.SpacingBeforePoints = double.PositiveInfinity);
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.LineSpacingPoints = -1.0);

        // w:hanging and w:firstLine are unsigned twips measures: a negative value would produce an
        // invalid package, so it is rejected instead of being written.
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.HangingIndentPoints = -1.0);

        // A value too large to be expressed in twips must fail loudly instead of overflowing into a
        // nonsense integer (which is what an unchecked (int) cast would produce).
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.LeftIndentPoints = double.MaxValue);
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.SpacingAfterPoints = double.MaxValue);
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.LineSpacingPoints = double.MaxValue);
    }

    [Fact]
    public void Paragraph_Alignment_CanBeChangedWithoutDuplicatingMarkup()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("x");

        paragraph.Alignment = ParagraphAlignment.Left;
        paragraph.Alignment = ParagraphAlignment.Center;

        Assert.Equal(ParagraphAlignment.Center, paragraph.Alignment);

        byte[] bytes = SaveBytes.ToBytes(document);
        string xml = TestPackage.ReadEntryText(bytes, "word/document.xml");
        Assert.Equal(1, Occurrences(xml, "<w:jc"));
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void Paragraph_UnmodelledStoredAlignment_ThrowsAndIsPreserved()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredAlignment("x", "mediumKashida");
        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Paragraph paragraph = document.Paragraphs[0];
            Assert.Throws<UnsupportedOfficeFeatureException>(() => paragraph.Alignment);

            // The unmodelled value must not be silently rewritten or dropped.
            paragraph.SpacingBeforePoints = 6.0;
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Contains("mediumKashida", TestPackage.ReadEntryText(produced, "word/document.xml"));
    }

    [Fact]
    public void Paragraph_MalformedStoredAlignment_ThrowsInvalidOfficeFile()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredAlignment("x", "bogus");
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].Alignment);
    }

    [Fact]
    public void Paragraph_StoredMultipleLineSpacing_ThrowsAndIsPreserved()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredLineSpacing("x", line: "360", lineRule: "auto");
        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Paragraph paragraph = document.Paragraphs[0];
            Assert.Throws<UnsupportedOfficeFeatureException>(() => paragraph.LineSpacingPoints);

            paragraph.SpacingBeforePoints = 6.0;
            produced = SaveBytes.ToBytes(document);
        }

        string xml = TestPackage.ReadEntryText(produced, "word/document.xml");
        Assert.Contains("w:lineRule=\"auto\"", xml);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_Indents_KeepUnmodelledAttributes()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredIndent("x", leftChars: 100, firstLine: "240");
        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Paragraph paragraph = document.Paragraphs[0];
            Assert.Equal(12.0, paragraph.FirstLineIndentPoints);

            paragraph.LeftIndentPoints = 10.0;
            produced = SaveBytes.ToBytes(document);
        }

        string xml = TestPackage.ReadEntryText(produced, "word/document.xml");
        Assert.Contains("w:leftChars=\"100\"", xml);
        Assert.Contains("w:left=\"200\"", xml);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_ClearingFormatting_RemovesTheEmptyProperties()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("clean");
        paragraph.Alignment = ParagraphAlignment.Center;
        paragraph.FirstLineIndentPoints = 12.0;
        paragraph.SpacingBeforePoints = 6.0;

        paragraph.Alignment = null;
        paragraph.FirstLineIndentPoints = null;
        paragraph.SpacingBeforePoints = null;

        byte[] bytes = SaveBytes.ToBytes(document);
        string xml = TestPackage.ReadEntryText(bytes, "word/document.xml");
        Assert.DoesNotContain("<w:pPr", xml);
        Assert.DoesNotContain("<w:ind", xml);
        Assert.DoesNotContain("<w:spacing", xml);
        Assert.DoesNotContain("<w:jc", xml);
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    private static int Occurrences(string text, string value)
        => text.Split(value, StringSplitOptions.None).Length - 1;

    [Fact]
    public void Paragraph_ClearingFormattingOnACleanParagraph_ChangesNothing()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("clean");

        paragraph.Alignment = null;
        paragraph.FirstLineIndentPoints = null;
        paragraph.LeftIndentPoints = null;
        paragraph.RightIndentPoints = null;
        paragraph.HangingIndentPoints = null;
        paragraph.SpacingBeforePoints = null;
        paragraph.SpacingAfterPoints = null;
        paragraph.LineSpacingPoints = null;

        byte[] bytes = SaveBytes.ToBytes(document);
        Assert.DoesNotContain("<w:pPr", TestPackage.ReadEntryText(bytes, "word/document.xml"));
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void Paragraph_ClearingOneIndent_KeepsTheOthers()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("x");
        paragraph.FirstLineIndentPoints = 12.0;
        paragraph.LeftIndentPoints = 24.0;

        paragraph.FirstLineIndentPoints = null;

        byte[] bytes = SaveBytes.ToBytes(document);
        string xml = TestPackage.ReadEntryText(bytes, "word/document.xml");
        Assert.DoesNotContain("w:firstLine", xml, StringComparison.Ordinal);
        Assert.Contains("w:left=\"480\"", xml, StringComparison.Ordinal);
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void Paragraph_RejectsUndefinedAlignment()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("x");

        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.Alignment = (ParagraphAlignment)99);
    }

    [Fact]
    public void Paragraph_MalformedStoredIndent_ThrowsInvalidOfficeFile()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredIndent("x", leftChars: null, firstLine: "bogus");
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].FirstLineIndentPoints);
    }

    [Fact]
    public void Paragraph_MalformedStoredLineSpacing_ThrowsInvalidOfficeFile()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredLineSpacing("x", line: "bogus", lineRule: "exact");
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].LineSpacingPoints);
    }

    [Fact]
    public void Paragraph_MalformedStoredLineSpacingRule_ThrowsInvalidOfficeFile()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredLineSpacing("x", line: "360", lineRule: "bogus");
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].LineSpacingPoints);
    }

    [Fact]
    public void Paragraph_MalformedStoredSpacing_ThrowsInvalidOfficeFile()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithStoredSpacing("x", before: "bogus");
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].SpacingBeforePoints);
    }
}
