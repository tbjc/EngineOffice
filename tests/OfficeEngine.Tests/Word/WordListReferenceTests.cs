using System;
using DocumentFormat.OpenXml;
using OfficeEngine;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.4: the paragraph list reference (<c>w:numPr/w:numId</c> and <c>w:numPr/w:ilvl</c>).
/// The reference is tri-state (<see langword="null"/> means "not named", <c>0</c> means "not
/// numbered") and is validated before it is written; the list it names is never created or modified
/// here.
/// </summary>
public sealed class WordListReferenceTests
{
    [Fact]
    public void Paragraph_ListId_IsNull_ByDefault()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Null(document.AddParagraph("text").ListId);
    }

    [Fact]
    public void Paragraph_ListLevel_IsNull_ByDefault()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Null(document.AddParagraph("text").ListLevel);
    }

    [Fact]
    public void Paragraph_ListId_ReadsAStoredReference()
    {
        using WordDocument document = OpenWithNumbering("text", 1, 0);

        Assert.Equal(1, document.Paragraphs[0].ListId);
    }

    [Fact]
    public void Paragraph_ListLevel_ReadsAStoredValue()
    {
        using WordDocument document = OpenWithNumbering("text", 4, 2);

        Assert.Equal(2, document.Paragraphs[0].ListLevel);
    }

    [Fact]
    public void Paragraph_ListId_ReadsTheReservedZero()
    {
        // w:numId="0" means "this paragraph is explicitly not numbered"; it is not "no reference".
        using WordDocument document = OpenWithNumbering("text", 0, null);

        Assert.Equal(0, document.Paragraphs[0].ListId);
    }

    [Fact]
    public void Paragraph_ListId_Throws_WhenTheValAttributeIsMissing()
    {
        byte[] package = OoxmlValidation.CreateDocxWithNumberingProperties(
            "text",
            numbering => numbering.AppendChild(new W.NumberingId()));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].ListId);
    }

    [Fact]
    public void Paragraph_ListLevel_Throws_WhenTheValAttributeIsMissing()
    {
        byte[] package = OoxmlValidation.CreateDocxWithNumberingProperties(
            "text",
            numbering => numbering.AppendChild(new W.NumberingLevelReference()));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].ListLevel);
    }

    [Fact]
    public void Paragraph_ListId_Throws_WhenTheStoredValueCannotBeParsed()
    {
        byte[] package = OoxmlValidation.CreateDocxWithNumberingProperties(
            "text",
            numbering => numbering.AppendChild(OoxmlValidation.WithAttribute(new W.NumberingId(), "val", "not-a-number")));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].ListId);
    }

    [Fact]
    public void Paragraph_ListLevel_Throws_WhenTheStoredValueCannotBeParsed()
    {
        byte[] package = OoxmlValidation.CreateDocxWithNumberingProperties(
            "text",
            numbering => numbering.AppendChild(OoxmlValidation.WithAttribute(new W.NumberingLevelReference(), "val", "x")));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].ListLevel);
    }

    [Fact]
    public void Paragraph_ListLevel_ReadsAValueOutsideTheRangeThisVersionWrites()
    {
        // The schema allows any decimal for w:ilvl; the validator accepts 12 and Word ignores it. It is
        // reported as stored rather than refused, because it is not this property's job to fix it.
        byte[] package = OoxmlValidation.CreateDocxWithNumberingProperties(
            "text",
            numbering => numbering.AppendChild(new W.NumberingLevelReference { Val = 12 }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        Assert.Equal(12, document.Paragraphs[0].ListLevel);
    }

    [Fact]
    public void Paragraph_ListId_IsReadableByTheSdkAfterSetting()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("text").ListId = 3;

        byte[] bytes = SaveBytes.ToBytes(document);

        W.NumberingProperties? numbering = OoxmlValidation.ReadDocxFirstParagraphProperties(bytes)
            ?.GetFirstChild<W.NumberingProperties>();
        Assert.Equal(3, numbering?.GetFirstChild<W.NumberingId>()?.Val?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void Paragraph_ListId_IsInsertedInSchemaOrder()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        // Applied in reverse schema order on purpose, so producing valid output depends on the
        // insertion logic rather than on the order of the assignments.
        paragraph.Alignment = ParagraphAlignment.Center;
        paragraph.LeftIndentPoints = 18;
        paragraph.AddTabStop(new TabStop(20, TabStopAlignment.Left));
        paragraph.ListId = 3;
        byte[] bytes = SaveBytes.ToBytes(document);

        W.ParagraphProperties? properties = OoxmlValidation.ReadDocxFirstParagraphProperties(bytes);
        Assert.NotNull(properties);
        Assert.IsType<W.NumberingProperties>(properties.ChildElements[0]);
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void Paragraph_ListId_AndLevel_AreInSchemaOrderWithinTheNumberingProperties()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        // Set the identifier first, so producing a valid order depends on the insertion logic.
        paragraph.ListId = 3;
        paragraph.ListLevel = 1;
        byte[] bytes = SaveBytes.ToBytes(document);

        W.NumberingProperties? numbering = OoxmlValidation.ReadDocxFirstParagraphProperties(bytes)
            ?.GetFirstChild<W.NumberingProperties>();
        Assert.NotNull(numbering);
        Assert.Equal("ilvl", numbering.ChildElements[0].LocalName);
        Assert.Equal("numId", numbering.ChildElements[1].LocalName);
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void Paragraph_ListId_OverwritesAnExistingReference()
    {
        using WordDocument document = OpenWithNumbering("text", 1, 0);
        Paragraph paragraph = document.Paragraphs[0];

        paragraph.ListId = 2;
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Equal(2, paragraph.ListId);
        Assert.Single(OoxmlValidation.ReadDocxFirstParagraphProperties(produced)!
            .GetFirstChild<W.NumberingProperties>()!.Elements<W.NumberingId>());
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_ListId_Null_ClearsTheReference()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        paragraph.ListId = 3;

        paragraph.ListId = null;
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Null(paragraph.ListId);
        Assert.Null(OoxmlValidation.ReadDocxFirstParagraphProperties(produced));
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_ListId_Null_KeepsTheLevel()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        paragraph.ListId = 3;
        paragraph.ListLevel = 2;

        paragraph.ListId = null;

        Assert.Null(paragraph.ListId);
        Assert.Equal(2, paragraph.ListLevel);
    }

    [Fact]
    public void Paragraph_ListLevel_Null_ClearsTheLevel()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        paragraph.ListLevel = 2;

        paragraph.ListLevel = null;
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Null(paragraph.ListLevel);
        Assert.Null(OoxmlValidation.ReadDocxFirstParagraphProperties(produced));
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_ListId_Null_IsANoOpWithoutAReference()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        paragraph.ListId = null;
        paragraph.ListLevel = null;

        Assert.Null(paragraph.ListId);
        Assert.Null(OoxmlValidation.ReadDocxFirstParagraphProperties(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void Paragraph_ListId_Null_KeepsOtherProperties()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        paragraph.ListId = 3;
        paragraph.Alignment = ParagraphAlignment.Right;

        paragraph.ListId = null;

        Assert.Null(paragraph.ListId);
        Assert.Equal(ParagraphAlignment.Right, paragraph.Alignment);
    }

    [Fact]
    public void Paragraph_ListId_RejectsANegativeValue()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.ListId = -1);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(9)]
    [InlineData(int.MaxValue)]
    public void Paragraph_ListLevel_RejectsAValueOutsideTheRange(int value)
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.ListLevel = value);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(4)]
    [InlineData(8)]
    public void Paragraph_ListLevel_AcceptsEveryValueInTheRange(int value)
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        paragraph.ListLevel = value;
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Equal(value, paragraph.ListLevel);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_ListId_ARejectedValueLeavesTheMarkupUntouched()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        paragraph.ListId = 3;

        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.ListId = -1);

        Assert.Equal(3, paragraph.ListId);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void Paragraph_ListId_PreservesUnmodelledParagraphProperties()
    {
        byte[] package = OoxmlValidation.CreateDocxWithUnmodelledParagraphFlags("text");
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));
        Paragraph paragraph = document.Paragraphs[0];

        paragraph.ListId = 3;
        byte[] produced = SaveBytes.ToBytes(document);

        W.ParagraphProperties? properties = OoxmlValidation.ReadDocxFirstParagraphProperties(produced);
        Assert.NotNull(properties);
        Assert.IsType<W.NumberingProperties>(properties.ChildElements[0]);
        Assert.NotNull(properties.GetFirstChild<W.OverflowPunctuation>());
        Assert.NotNull(properties.GetFirstChild<W.BiDi>());
        Assert.Equal(3, paragraph.ListId);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_ListId_RoundTripsThroughSaveAndOpen()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            Paragraph paragraph = document.AddParagraph("text");
            paragraph.ListId = 5;
            paragraph.ListLevel = 1;
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Open(stream);
        Assert.Equal(5, reopened.Paragraphs[0].ListId);
        Assert.Equal(1, reopened.Paragraphs[0].ListLevel);
    }

    [Fact]
    public void Paragraph_ListId_RoundTripsThroughAFile()
    {
        string path = Path.Combine(Path.GetTempPath(), $"oe-list-{Guid.NewGuid():N}.docx");
        try
        {
            using (WordDocument document = WordDocument.Create())
            {
                document.AddParagraph("text").ListId = 5;
                document.Save(path);
            }

            using WordDocument reopened = WordDocument.Open(path);
            Assert.Equal(5, reopened.Paragraphs[0].ListId);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Paragraph_ApplyList_SetsTheIdentifierAndTheLevel()
    {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListFormat.Decimal);
        Paragraph paragraph = document.AddParagraph("text");

        paragraph.ApplyList(list);
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Equal(list.Id, paragraph.ListId);
        Assert.Equal(0, paragraph.ListLevel);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_ApplyList_UsesTheRequestedLevel()
    {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListFormat.Bullet);
        Paragraph paragraph = document.AddParagraph("text");

        paragraph.ApplyList(list, 3);

        Assert.Equal(list.Id, paragraph.ListId);
        Assert.Equal(3, paragraph.ListLevel);
    }

    [Fact]
    public void Paragraph_ApplyList_RejectsNull()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        Assert.Throws<ArgumentNullException>(() => paragraph.ApplyList(null!));
    }

    [Fact]
    public void Paragraph_ApplyList_RejectsALevelOutsideTheRange()
    {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListFormat.Decimal);
        Paragraph paragraph = document.AddParagraph("text");

        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.ApplyList(list, 9));
    }

    private static WordDocument OpenWithNumbering(string text, int? listId, int? level)
    {
        return WordDocument.Open(TestPackage.ToStream(
            OoxmlValidation.CreateDocxWithNumberingProperties(text, numbering =>
            {
                if (level is not null)
                {
                    numbering.AppendChild(new W.NumberingLevelReference { Val = level.Value });
                }

                if (listId is not null)
                {
                    numbering.AppendChild(new W.NumberingId { Val = listId.Value });
                }
            })));
    }
}
