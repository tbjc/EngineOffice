using System;
using DocumentFormat.OpenXml;
using OfficeEngine;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.4: the numbering catalogue (<c>word/numbering.xml</c>), the levels it exposes, and the
/// list definitions <see cref="WordDocument.AddList"/> creates.
/// </summary>
public sealed class WordListCatalogTests
{
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly string[] TwoListsChildOrder = ["abstractNum", "abstractNum", "num", "num"];

    [Fact]
    public void Lists_IsEmpty_WhenTheDocumentHasNoNumberingPart()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Empty(document.Lists);
        Assert.Null(document.FindList(1));
    }

    [Fact]
    public void Lists_IsEmpty_WhenTheNumberingPartHasNoCatalogue()
    {
        byte[] package = OoxmlValidation.CreateDocxWithEmptyNumberingPart("body");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        Assert.Empty(document.Lists);
        Assert.Null(document.FindList(1));
    }

    [Fact]
    public void Lists_ReturnsTheDefinitionsInDocumentOrder()
    {
        using WordDocument document = OpenWithLists(
            OoxmlValidation.AbstractNumbering(0, W.NumberFormatValues.Bullet),
            OoxmlValidation.AbstractNumbering(1, W.NumberFormatValues.Decimal),
            OoxmlValidation.ListInstance(1, 0),
            OoxmlValidation.ListInstance(2, 1));

        Assert.Equal(2, document.Lists.Count);
        Assert.Equal(1, document.Lists[0].Id);
        Assert.Equal(2, document.Lists[1].Id);
    }

    [Fact]
    public void List_ExposesItsIdentity()
    {
        using WordDocument document = OpenWithLists(
            OoxmlValidation.AbstractNumbering(7, W.NumberFormatValues.Decimal),
            OoxmlValidation.ListInstance(3, 7));

        WordList list = document.Lists[0];

        Assert.Equal(3, list.Id);
        Assert.Equal(7, list.AbstractNumberId);
        Assert.Equal("3", list.ToString());
        Assert.Equal(3, document.FindList(3)!.Id);
    }

    [Fact]
    public void List_Levels_AreReadFromTheReferencedAbstractNumbering()
    {
        W.AbstractNum definition = new(
            new W.Level { LevelIndex = 0, NumberingFormat = new W.NumberingFormat { Val = W.NumberFormatValues.Decimal } },
            new W.Level { LevelIndex = 1, NumberingFormat = new W.NumberingFormat { Val = W.NumberFormatValues.LowerLetter } })
        { AbstractNumberId = 0 };

        using WordDocument document = OpenWithLists(definition, OoxmlValidation.ListInstance(1, 0));

        Assert.Equal(2, document.Lists[0].Levels.Count);
        Assert.Equal(0, document.Lists[0].Levels[0].Level);
        Assert.Equal(1, document.Lists[0].Levels[1].Level);
        Assert.Equal(WordListFormat.Decimal, document.Lists[0].Levels[0].Format);
        Assert.Equal(WordListFormat.LowerLetter, document.Lists[0].Levels[1].Format);
    }

    [Fact]
    public void List_Levels_IsEmpty_WhenTheAbstractNumberingIsMissing()
    {
        // A w:num that references an abstract definition the part does not contain is valid; Word
        // renders its paragraphs as if they were not numbered.
        using WordDocument document = OpenWithLists(OoxmlValidation.ListInstance(1, 9));

        WordList list = document.Lists[0];

        Assert.Equal(9, list.AbstractNumberId);
        Assert.Empty(list.Levels);
        Assert.Null(list.FindLevel(0));
    }

    [Fact]
    public void List_Level_ExposesItsIdentity()
    {
        using WordDocument document = OpenWithLists(
            OoxmlValidation.AbstractNumbering(0, W.NumberFormatValues.UpperRoman),
            OoxmlValidation.ListInstance(1, 0));

        WordListLevel level = document.Lists[0].Levels[0];

        Assert.Equal(0, level.Level);
        Assert.Equal(1, level.Start);
        Assert.Equal(WordListFormat.UpperRoman, level.Format);
        Assert.Equal("%1.", level.Text);
        Assert.Equal(WordListJustification.Left, level.Justification);
    }

    [Fact]
    public void List_FindLevel_ReturnsTheMatchingLevel()
    {
        using WordDocument document = OpenWithLists(
            AbstractWithLevels(0, 0, 1),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Equal(1, document.Lists[0].FindLevel(1)!.Level);
    }

    [Fact]
    public void List_FindLevel_ReturnsNull_WhenTheLevelIsAbsent()
    {
        using WordDocument document = OpenWithLists(
            OoxmlValidation.AbstractNumbering(0, W.NumberFormatValues.Decimal),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Null(document.Lists[0].FindLevel(5));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(9)]
    public void List_FindLevel_RejectsALevelOutsideTheRange(int level)
    {
        using WordDocument document = OpenWithLists(
            OoxmlValidation.AbstractNumbering(0, W.NumberFormatValues.Decimal),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Throws<ArgumentOutOfRangeException>(() => document.Lists[0].FindLevel(level));
    }

    [Fact]
    public void List_Level_Format_MapsEveryModelledFormat()
    {
        (W.NumberFormatValues Stored, WordListFormat Expected)[] cases =
        [
            (W.NumberFormatValues.Bullet, WordListFormat.Bullet),
            (W.NumberFormatValues.Decimal, WordListFormat.Decimal),
            (W.NumberFormatValues.LowerLetter, WordListFormat.LowerLetter),
            (W.NumberFormatValues.UpperLetter, WordListFormat.UpperLetter),
            (W.NumberFormatValues.LowerRoman, WordListFormat.LowerRoman),
            (W.NumberFormatValues.UpperRoman, WordListFormat.UpperRoman),
        ];

        foreach ((W.NumberFormatValues stored, WordListFormat expected) in cases)
        {
            using WordDocument document = OpenWithLists(
                OoxmlValidation.AbstractNumbering(0, stored),
                OoxmlValidation.ListInstance(1, 0));

            Assert.Equal(expected, document.Lists[0].Levels[0].Format);
        }
    }

    [Fact]
    public void List_Level_Format_Throws_WhenTheValueIsNotModelled()
    {
        // 'chicago' is a valid ST_NumberFormat that this version does not model.
        using WordDocument document = OpenWithLists(
            AbstractWithLevel(0, OoxmlValidation.WithAttribute(new W.NumberingFormat(), "val", "chicago")),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Throws<UnsupportedOfficeFeatureException>(() => document.Lists[0].Levels[0].Format);
    }

    [Fact]
    public void List_Level_Format_Throws_WhenTheStoredValueCannotBeParsed()
    {
        using WordDocument document = OpenWithLists(
            AbstractWithLevel(0, OoxmlValidation.WithAttribute(new W.NumberingFormat(), "val", "bogus")),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Throws<InvalidOfficeFileException>(() => document.Lists[0].Levels[0].Format);
    }

    [Fact]
    public void List_Level_Format_Throws_WhenTheValAttributeIsMissing()
    {
        using WordDocument document = OpenWithLists(
            AbstractWithLevel(0, new W.NumberingFormat()),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Throws<InvalidOfficeFileException>(() => document.Lists[0].Levels[0].Format);
    }

    [Fact]
    public void List_Level_Format_IsNull_WhenTheLevelDoesNotDeclareOne()
    {
        using WordDocument document = OpenWithLists(
            AbstractWithLevel(0, new W.LevelText { Val = "%1." }),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Null(document.Lists[0].Levels[0].Format);
    }

    [Fact]
    public void List_Level_Start_IsNull_WhenTheLevelDoesNotSetOne()
    {
        using WordDocument document = OpenWithLists(
            AbstractWithLevel(0, new W.NumberingFormat { Val = W.NumberFormatValues.Decimal }),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Null(document.Lists[0].Levels[0].Start);
    }

    [Fact]
    public void List_Level_Start_Throws_WhenTheValAttributeIsMissing()
    {
        using WordDocument document = OpenWithLists(
            AbstractWithLevel(0, new W.StartNumberingValue()),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Throws<InvalidOfficeFileException>(() => document.Lists[0].Levels[0].Start);
    }

    [Fact]
    public void List_Level_Text_IsNull_WhenTheLevelDoesNotSetOne()
    {
        using WordDocument document = OpenWithLists(
            AbstractWithLevel(0, new W.NumberingFormat { Val = W.NumberFormatValues.Decimal }),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Null(document.Lists[0].Levels[0].Text);
    }

    [Fact]
    public void List_Level_Text_Throws_WhenTheValAttributeIsMissing()
    {
        using WordDocument document = OpenWithLists(
            AbstractWithLevel(0, new W.LevelText()),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Throws<InvalidOfficeFileException>(() => document.Lists[0].Levels[0].Text);
    }

    [Fact]
    public void List_Level_Justification_MapsEveryValue()
    {
        (W.LevelJustificationValues Stored, WordListJustification Expected)[] cases =
        [
            (W.LevelJustificationValues.Left, WordListJustification.Left),
            (W.LevelJustificationValues.Center, WordListJustification.Center),
            (W.LevelJustificationValues.Right, WordListJustification.Right),
        ];

        foreach ((W.LevelJustificationValues stored, WordListJustification expected) in cases)
        {
            using WordDocument document = OpenWithLists(
                AbstractWithLevel(0, new W.LevelJustification { Val = stored }),
                OoxmlValidation.ListInstance(1, 0));

            Assert.Equal(expected, document.Lists[0].Levels[0].Justification);
        }
    }

    [Fact]
    public void List_Level_Justification_IsNull_WhenTheLevelDoesNotSetOne()
    {
        using WordDocument document = OpenWithLists(
            AbstractWithLevel(0, new W.NumberingFormat { Val = W.NumberFormatValues.Decimal }),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Null(document.Lists[0].Levels[0].Justification);
    }

    [Fact]
    public void List_Level_Justification_Throws_WhenTheValAttributeIsMissing()
    {
        using WordDocument document = OpenWithLists(
            AbstractWithLevel(0, new W.LevelJustification()),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Throws<InvalidOfficeFileException>(() => document.Lists[0].Levels[0].Justification);
    }

    [Fact]
    public void List_Level_Justification_Throws_WhenTheStoredValueCannotBeParsed()
    {
        using WordDocument document = OpenWithLists(
            AbstractWithLevel(0, OoxmlValidation.WithAttribute(new W.LevelJustification(), "val", "both")),
            OoxmlValidation.ListInstance(1, 0));

        Assert.Throws<InvalidOfficeFileException>(() => document.Lists[0].Levels[0].Justification);
    }

    [Fact]
    public void List_Level_Throws_WhenTheIndexIsMissing()
    {
        using WordDocument document = OpenWithLists(
            new W.AbstractNum(new W.Level()) { AbstractNumberId = 0 },
            OoxmlValidation.ListInstance(1, 0));

        Assert.Throws<InvalidOfficeFileException>(() => document.Lists[0].Levels[0].Level);
    }

    [Fact]
    public void List_Id_Throws_WhenTheIdentifierIsMissing()
    {
        using WordDocument document = OpenWithLists(new W.NumberingInstance(new W.AbstractNumId { Val = 0 }));

        WordList list = document.Lists[0];

        Assert.Throws<InvalidOfficeFileException>(() => list.Id);
        Assert.Equal("(list without an id)", list.ToString());
    }

    [Fact]
    public void List_Id_Throws_WhenTheIdentifierCannotBeParsed()
    {
        W.NumberingInstance instance = new(new W.AbstractNumId { Val = 0 });
        instance.SetAttribute(new OpenXmlAttribute("w", "numId", WordprocessingNamespace, "nope"));

        using WordDocument document = OpenWithLists(instance);

        Assert.Throws<InvalidOfficeFileException>(() => document.Lists[0].Id);
    }

    [Fact]
    public void List_AbstractNumberId_Throws_WhenTheReferenceIsMissing()
    {
        using WordDocument document = OpenWithLists(new W.NumberingInstance { NumberID = 1 });

        Assert.Throws<InvalidOfficeFileException>(() => document.Lists[0].AbstractNumberId);
    }

    [Fact]
    public void List_AbstractNumberId_Throws_WhenTheValAttributeIsMissing()
    {
        using WordDocument document = OpenWithLists(new W.NumberingInstance(new W.AbstractNumId()) { NumberID = 1 });

        Assert.Throws<InvalidOfficeFileException>(() => document.Lists[0].AbstractNumberId);
    }

    [Fact]
    public void FindList_SkipsADefinitionWhoseIdentifierCannotBeRead()
    {
        W.NumberingInstance malformed = new(new W.AbstractNumId { Val = 0 });
        malformed.SetAttribute(new OpenXmlAttribute("w", "numId", WordprocessingNamespace, "nope"));

        using WordDocument document = OpenWithLists(
            OoxmlValidation.AbstractNumbering(0, W.NumberFormatValues.Decimal),
            malformed,
            OoxmlValidation.ListInstance(2, 0));

        Assert.Equal(2, document.FindList(2)!.Id);
        Assert.Null(document.FindList(1));
    }

    [Fact]
    public void AddList_CreatesANumberingPartAndReturnsAUsableDefinition()
    {
        using WordDocument document = WordDocument.Create();

        WordList list = document.AddList(WordListFormat.Decimal);

        Assert.Equal(1, list.Id);
        Assert.Equal(0, list.AbstractNumberId);
        Assert.Single(document.Lists);
        Assert.Equal(1, document.FindList(1)!.Id);
    }

    [Fact]
    public void AddList_AppendsANumberedLevelInSchemaOrder()
    {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListFormat.Decimal);
        byte[] bytes = SaveBytes.ToBytes(document);

        W.AbstractNum abstractNumbering = OoxmlValidation.ReadDocxNumbering(bytes)!
            .Elements<W.AbstractNum>().Single();
        W.Level level = abstractNumbering.Elements<W.Level>().Single();

        Assert.Equal(0, abstractNumbering.AbstractNumberId?.Value);
        Assert.Equal(W.MultiLevelValues.SingleLevel, abstractNumbering.MultiLevelType?.Val?.Value);
        Assert.Equal(0, level.LevelIndex?.Value);
        Assert.Equal(1, level.StartNumberingValue?.Val?.Value);
        Assert.Equal(W.NumberFormatValues.Decimal, level.NumberingFormat?.Val?.Value);
        Assert.Equal("%1.", level.LevelText?.Val?.Value);
        Assert.Equal(W.LevelJustificationValues.Left, level.LevelJustification?.Val?.Value);
        Assert.Equal("720", level.PreviousParagraphProperties?.Indentation?.Left?.Value);
        Assert.Equal("360", level.PreviousParagraphProperties?.Indentation?.Hanging?.Value);
        Assert.Null(level.NumberingSymbolRunProperties);
        Assert.Equal(WordListFormat.Decimal, list.Levels[0].Format);
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void AddList_AppendsABulletLevelWithTheSymbolFont()
    {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListFormat.Bullet);
        byte[] bytes = SaveBytes.ToBytes(document);

        W.Level level = OoxmlValidation.ReadDocxNumbering(bytes)!
            .Elements<W.AbstractNum>().Single().Elements<W.Level>().Single();

        Assert.Equal(W.NumberFormatValues.Bullet, level.NumberingFormat?.Val?.Value);
        Assert.Equal("\uF0B7", level.LevelText?.Val?.Value);
        Assert.Equal("Symbol", level.NumberingSymbolRunProperties?.RunFonts?.Ascii?.Value);
        Assert.Equal("Symbol", level.NumberingSymbolRunProperties?.RunFonts?.HighAnsi?.Value);
        Assert.Equal(WordListFormat.Bullet, list.Levels[0].Format);
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void AddList_WritesEveryAbstractNumberingBeforeTheInstances()
    {
        using WordDocument document = WordDocument.Create();

        document.AddList(WordListFormat.Decimal);
        document.AddList(WordListFormat.Bullet);
        byte[] bytes = SaveBytes.ToBytes(document);

        W.Numbering root = OoxmlValidation.ReadDocxNumbering(bytes)!;
        Assert.Equal(
            TwoListsChildOrder,
            root.ChildElements.Select(element => element.LocalName));
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void AddList_WritesEveryModelledFormat()
    {
        WordListFormat[] formats =
        [
            WordListFormat.Bullet,
            WordListFormat.Decimal,
            WordListFormat.LowerLetter,
            WordListFormat.UpperLetter,
            WordListFormat.LowerRoman,
            WordListFormat.UpperRoman,
        ];

        using WordDocument document = WordDocument.Create();
        foreach (WordListFormat format in formats)
        {
            WordList list = document.AddList(format);
            Assert.Equal(format, list.Levels[0].Format);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void AddList_AssignsIncrementingIdentifiers()
    {
        using WordDocument document = WordDocument.Create();

        WordList first = document.AddList(WordListFormat.Decimal);
        WordList second = document.AddList(WordListFormat.Bullet);

        Assert.Equal(1, first.Id);
        Assert.Equal(2, second.Id);
        Assert.Equal(0, first.AbstractNumberId);
        Assert.Equal(1, second.AbstractNumberId);
    }

    [Fact]
    public void AddList_PlacesItsDefinitionsBeforeExistingInstances()
    {
        byte[] package = OoxmlValidation.CreateDocxWithLists(
            "text",
            null,
            OoxmlValidation.AbstractNumbering(0, W.NumberFormatValues.Bullet),
            OoxmlValidation.ListInstance(1, 0));
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        WordList created = document.AddList(WordListFormat.Decimal);
        byte[] bytes = SaveBytes.ToBytes(document);

        W.Numbering root = OoxmlValidation.ReadDocxNumbering(bytes)!;
        Assert.Equal(2, created.Id);
        Assert.Equal(
            TwoListsChildOrder,
            root.ChildElements.Select(element => element.LocalName));
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void AddList_RejectsAnUndefinedFormatWithoutTouchingTheDocument()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddList((WordListFormat)42));

        Assert.Empty(document.Lists);
        Assert.Null(OoxmlValidation.ReadDocxNumbering(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void AddList_SurvivesSaveAndOpen()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            WordList list = document.AddList(WordListFormat.LowerRoman);
            document.AddParagraph("primero").ApplyList(list);
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Open(stream);
        WordList restored = Assert.Single(reopened.Lists);

        Assert.Equal(1, restored.Id);
        Assert.Equal(WordListFormat.LowerRoman, restored.Levels[0].Format);
        Assert.Equal(1, reopened.Paragraphs[0].ListId);
        Assert.Equal(0, reopened.Paragraphs[0].ListLevel);
    }

    private static WordDocument OpenWithLists(params OpenXmlElement[] numberingChildren)
    {
        return WordDocument.Open(TestPackage.ToStream(
            OoxmlValidation.CreateDocxWithLists("text", null, numberingChildren)));
    }

    private static W.AbstractNum AbstractWithLevel(int abstractNumberId, params OpenXmlElement[] levelChildren)
    {
        W.Level level = new() { LevelIndex = 0 };
        foreach (OpenXmlElement child in levelChildren)
        {
            level.AppendChild(child);
        }

        return new W.AbstractNum(level) { AbstractNumberId = abstractNumberId };
    }

    private static W.AbstractNum AbstractWithLevels(int abstractNumberId, int firstLevel, int secondLevel)
    {
        return new W.AbstractNum(
            new W.Level { LevelIndex = firstLevel },
            new W.Level { LevelIndex = secondLevel })
        { AbstractNumberId = abstractNumberId };
    }
}
