using System;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.4 interoperability checks. Every assertion is made against the Open XML SDK or against
/// the raw package content, never through the code under test.
/// </summary>
public sealed class WordListInteropTests
{
    [Fact]
    public void ProducedDocument_WithLists_IsSchemaValid()
    {
        using WordDocument document = WordDocument.Create();
        WordList bullets = document.AddList(WordListFormat.Bullet);
        WordList numbers = document.AddList(WordListFormat.Decimal);
        document.AddParagraph("bullet").ApplyList(bullets);
        document.AddParagraph("number").ApplyList(numbers, 0);
        byte[] bytes = SaveBytes.ToBytes(document);

        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void ProducedDocument_ListReferences_AreReadableByTheOpenXmlSdk()
    {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListFormat.LowerLetter);
        document.AddParagraph("first").ApplyList(list);
        document.AddParagraph("second").ApplyList(list, 0);
        byte[] bytes = SaveBytes.ToBytes(document);

        W.NumberingProperties? numbering = OoxmlValidation.ReadDocxFirstParagraphProperties(bytes)
            ?.GetFirstChild<W.NumberingProperties>();
        Assert.Equal(list.Id, numbering?.GetFirstChild<W.NumberingId>()?.Val?.Value);
        Assert.Equal(0, numbering?.GetFirstChild<W.NumberingLevelReference>()?.Val?.Value);
    }

    [Fact]
    public void ProducedDocument_DoesNotInventANumberingPart_WhenNoListIsCreated()
    {
        using WordDocument document = WordDocument.Create();

        // A dangling reference is valid and does not require a numbering part.
        document.AddParagraph("text").ListId = 3;
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.False(TestPackage.HasEntry(produced, "word/numbering.xml"));
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void ForeignDocument_ListLevelsAreReadFromTheRawMarkup()
    {
        W.AbstractNum definition = new(
            new W.Level
            {
                LevelIndex = 0,
                StartNumberingValue = new W.StartNumberingValue { Val = 1 },
                NumberingFormat = new W.NumberingFormat { Val = W.NumberFormatValues.UpperRoman },
                LevelText = new W.LevelText { Val = "%1)" },
                LevelJustification = new W.LevelJustification { Val = W.LevelJustificationValues.Right },
            },
            new W.Level
            {
                LevelIndex = 1,
                NumberingFormat = new W.NumberingFormat { Val = W.NumberFormatValues.LowerLetter },
                LevelText = new W.LevelText { Val = "%2." },
            })
        { AbstractNumberId = 4 };
        byte[] package = OoxmlValidation.CreateDocxWithLists(
            "text",
            new W.NumberingProperties(new W.NumberingLevelReference { Val = 1 }, new W.NumberingId { Val = 2 }),
            definition,
            OoxmlValidation.ListInstance(2, 4));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));
        WordList list = document.FindList(2)!;

        Assert.Equal(4, list.AbstractNumberId);
        Assert.Equal(2, list.Levels.Count);
        Assert.Equal(WordListFormat.UpperRoman, list.Levels[0].Format);
        Assert.Equal("%1)", list.Levels[0].Text);
        Assert.Equal(WordListJustification.Right, list.Levels[0].Justification);
        Assert.Equal(WordListFormat.LowerLetter, list.FindLevel(1)!.Format);
        Assert.Equal(1, document.Paragraphs[0].ListLevel);
        Assert.Equal(2, document.Paragraphs[0].ListId);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void NumberingPart_IsPreservedWhenTheBodyIsEdited()
    {
        W.AbstractNum definition = new(
            new W.Level
            {
                LevelIndex = 0,
                StartNumberingValue = new W.StartNumberingValue { Val = 1 },
                NumberingFormat = new W.NumberingFormat { Val = W.NumberFormatValues.Decimal },
                LevelText = new W.LevelText { Val = "%1." },
            })
        {
            AbstractNumberId = 0,
            Nsid = new W.Nsid { Val = "0A1B2C3D" },
        };
        byte[] package = OoxmlValidation.CreateDocxWithLists(
            "text", null, definition, OoxmlValidation.ListInstance(1, 0));
        string before = OoxmlValidation.ReadDocxNumberingXml(package);

        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(package)))
        {
            document.Paragraphs[0].Text = "edited";
            document.Save(stream);
        }

        byte[] produced = stream.ToArray();
        Assert.Equal(before, OoxmlValidation.ReadDocxNumberingXml(produced));
        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Equal(WordListFormat.Decimal, reopened.FindList(1)!.Levels[0].Format);
    }

    [Fact]
    public void ExistingNumberingIsPreserved_WhenANewListIsAdded()
    {
        W.AbstractNum foreign = new(
            new W.Level
            {
                LevelIndex = 0,
                StartNumberingValue = new W.StartNumberingValue { Val = 1 },
                NumberingFormat = new W.NumberingFormat { Val = W.NumberFormatValues.LowerRoman },
                LevelText = new W.LevelText { Val = "(%1)" },
            })
        {
            AbstractNumberId = 0,
            Nsid = new W.Nsid { Val = "0A1B2C3D" },
        };
        byte[] package = OoxmlValidation.CreateDocxWithLists(
            "text", null, foreign, OoxmlValidation.ListInstance(1, 0));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));
        document.AddList(WordListFormat.Decimal);
        byte[] produced = SaveBytes.ToBytes(document);

        string numberingXml = OoxmlValidation.ReadDocxNumberingXml(produced);
        Assert.Contains("0A1B2C3D", numberingXml, StringComparison.Ordinal);
        Assert.Contains("(%1)", numberingXml, StringComparison.Ordinal);
        Assert.Equal(WordListFormat.LowerRoman, document.FindList(1)!.Levels[0].Format);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void ParagraphFormattingEdits_DoNotBreakTheList()
    {
        byte[] package = OoxmlValidation.CreateDocxWithLists(
            "text",
            new W.NumberingProperties(new W.NumberingLevelReference { Val = 0 }, new W.NumberingId { Val = 1 }),
            OoxmlValidation.AbstractNumbering(0, W.NumberFormatValues.Decimal),
            OoxmlValidation.ListInstance(1, 0));
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));
        Paragraph paragraph = document.Paragraphs[0];

        paragraph.Alignment = ParagraphAlignment.Center;
        paragraph.SpacingAfterPoints = 12;
        byte[] produced = SaveBytes.ToBytes(document);

        W.NumberingProperties? numbering = OoxmlValidation.ReadDocxFirstParagraphProperties(produced)
            ?.GetFirstChild<W.NumberingProperties>();
        Assert.Equal(1, numbering?.GetFirstChild<W.NumberingId>()?.Val?.Value);
        Assert.Equal(0, numbering?.GetFirstChild<W.NumberingLevelReference>()?.Val?.Value);
        Assert.Equal(WordListFormat.Decimal, document.FindList(1)!.Levels[0].Format);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void DanglingListReference_RoundTrips()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("text").ListId = 99;
            document.Save(stream);
        }

        byte[] produced = stream.ToArray();
        W.Paragraph? paragraph = OoxmlValidation.ReadDocxFirstParagraph(produced);
        Assert.Equal(99, paragraph?.ParagraphProperties?.NumberingProperties
            ?.GetFirstChild<W.NumberingId>()?.Val?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void ListReferences_WrittenByTheLibrary_SurviveAFileRoundTrip()
    {
        string path = Path.Combine(Path.GetTempPath(), $"oe-list-interop-{Guid.NewGuid():N}.docx");
        try
        {
            WordList list;
            using (WordDocument document = WordDocument.Create())
            {
                list = document.AddList(WordListFormat.UpperLetter);
                document.AddParagraph("item").ApplyList(list, 2);
                document.Save(path);
            }

            using WordDocument reopened = WordDocument.Open(path);
            Assert.Equal(list.Id, reopened.Paragraphs[0].ListId);
            Assert.Equal(2, reopened.Paragraphs[0].ListLevel);
            Assert.Equal(WordListFormat.UpperLetter, reopened.FindList(list.Id)!.Levels[0].Format);
            Assert.Equal(W.NumberFormatValues.UpperLetter, OoxmlValidation.ReadDocxNumbering(File.ReadAllBytes(path))!
                .Elements<W.AbstractNum>().Single().Elements<W.Level>().Single().NumberingFormat?.Val?.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
