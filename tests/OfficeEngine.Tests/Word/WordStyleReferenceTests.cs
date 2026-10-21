using System;
using OfficeEngine;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.3: the paragraph and run style references (<c>w:pStyle</c>, <c>w:rStyle</c>). The
/// reference is tri-state and is validated before it is written; the style it names is never created
/// or modified.
/// </summary>
public sealed class WordStyleReferenceTests
{
    [Fact]
    public void Paragraph_StyleId_IsNull_ByDefault()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Null(document.AddParagraph("text").StyleId);
    }

    [Fact]
    public void Paragraph_StyleId_ReadsAStoredReference()
    {
        byte[] package = OoxmlValidation.CreateDocxWithStyleReference("text", "Heading1", null);

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        Assert.Equal("Heading1", document.Paragraphs[0].StyleId);
    }

    [Fact]
    public void Paragraph_StyleId_ReadsAStoredReference_WithoutAstylePart()
    {
        // A w:pStyle that names a style the document does not define is valid and is read as given.
        byte[] package = OoxmlValidation.CreateDocxWithStoredParagraphStyleReference("text", "MissingStyle");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        Assert.Equal("MissingStyle", document.Paragraphs[0].StyleId);
        Assert.Empty(OoxmlValidation.ValidateDocx(package));
    }

    [Fact]
    public void Paragraph_StyleId_Throws_WhenTheValAttributeIsMissing()
    {
        byte[] package = OoxmlValidation.CreateDocxWithStoredParagraphStyleReference("text", null);

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].StyleId);
    }

    [Fact]
    public void Paragraph_StyleId_IsReadableByTheSdkAfterSetting()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("text").StyleId = "Heading1";

        byte[] bytes = SaveBytes.ToBytes(document);

        W.ParagraphStyleId? style = OoxmlValidation.ReadDocxFirstParagraphProperties(bytes)
            ?.GetFirstChild<W.ParagraphStyleId>();
        Assert.Equal("Heading1", style?.Val?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void Paragraph_StyleId_IsInsertedInSchemaOrder()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        // Applied in reverse schema order on purpose, so producing valid output depends on the
        // insertion logic rather than on the order of the assignments.
        paragraph.Alignment = ParagraphAlignment.Center;
        paragraph.LeftIndentPoints = 18;
        paragraph.AddTabStop(new TabStop(20, TabStopAlignment.Left));
        paragraph.StyleId = "Heading1";
        byte[] bytes = SaveBytes.ToBytes(document);

        W.ParagraphProperties? properties = OoxmlValidation.ReadDocxFirstParagraphProperties(bytes);
        Assert.NotNull(properties);
        Assert.IsType<W.ParagraphStyleId>(properties.ChildElements[0]);
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void Paragraph_StyleId_OverwritesAnExistingReference()
    {
        byte[] package = OoxmlValidation.CreateDocxWithStyleReference("text", "Old", null);
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));
        Paragraph paragraph = document.Paragraphs[0];

        paragraph.StyleId = "New";
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Equal("New", paragraph.StyleId);
        Assert.Single(OoxmlValidation.ReadDocxFirstParagraphProperties(produced)!.Elements<W.ParagraphStyleId>());
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_StyleId_Null_ClearsTheReference()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        paragraph.StyleId = "Heading1";

        paragraph.StyleId = null;
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Null(paragraph.StyleId);
        Assert.Null(OoxmlValidation.ReadDocxFirstParagraphProperties(produced));
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_StyleId_Null_KeepsOtherProperties()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        paragraph.StyleId = "Heading1";
        paragraph.Alignment = ParagraphAlignment.Right;

        paragraph.StyleId = null;

        Assert.Null(paragraph.StyleId);
        Assert.Equal(ParagraphAlignment.Right, paragraph.Alignment);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("\t")]
    [InlineData("\n")]
    [InlineData("\r")]
    [InlineData("a\tb")]
    [InlineData("a\nb")]
    public void Paragraph_StyleId_RejectsAValueThatIsNotAnId(string value)
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        Assert.Throws<ArgumentException>(() => paragraph.StyleId = value);
    }

    [Fact]
    public void Paragraph_StyleId_RejectsAValueLongerThanTheSchemaAllows()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        Assert.Throws<ArgumentException>(() => paragraph.StyleId = new string('a', 254));
    }

    [Fact]
    public void Paragraph_StyleId_AcceptsAValueAtTheSchemaLimit()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        string value = new('a', 253);

        paragraph.StyleId = value;
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Equal(value, paragraph.StyleId);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_StyleId_RejectsACharacterXmlCannotStore()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        Assert.Throws<ArgumentException>(() => paragraph.StyleId = "bad\uD800");
    }

    [Fact]
    public void Paragraph_StyleId_AcceptsAValueWithAnAstralCharacter()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        paragraph.StyleId = "My\U0001F600Style";
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Equal("My\U0001F600Style", paragraph.StyleId);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_StyleId_AcceptsSpacesInsideTheIdentifier()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");

        paragraph.StyleId = "My Style";
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Equal("My Style", paragraph.StyleId);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_StyleId_ARejectedValueLeavesTheMarkupUntouched()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("text");
        paragraph.StyleId = "Heading1";

        Assert.Throws<ArgumentException>(() => paragraph.StyleId = " ");

        Assert.Equal("Heading1", paragraph.StyleId);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void Paragraph_StyleId_PreservesUnmodelledProperties()
    {
        byte[] package = OoxmlValidation.CreateDocxWithNumbering("text");
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));
        Paragraph paragraph = document.Paragraphs[0];

        paragraph.StyleId = "ListParagraph";
        byte[] produced = SaveBytes.ToBytes(document);

        string xml = TestPackage.ReadEntryText(produced, "word/document.xml");
        Assert.Contains("w:numPr", xml, StringComparison.Ordinal);
        Assert.Equal("ListParagraph", paragraph.StyleId);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Paragraph_StyleId_RoundTripsThroughSaveAndOpen()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("text").StyleId = "Heading1";
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Open(stream);
        Assert.Equal("Heading1", reopened.Paragraphs[0].StyleId);
    }

    [Fact]
    public void Paragraph_StyleId_RoundTripsThroughAFile()
    {
        string path = Path.Combine(Path.GetTempPath(), $"oe-style-{Guid.NewGuid():N}.docx");
        try
        {
            using (WordDocument document = WordDocument.Create())
            {
                document.AddParagraph("text").StyleId = "Heading1";
                document.Save(path);
            }

            using WordDocument reopened = WordDocument.Open(path);
            Assert.Equal("Heading1", reopened.Paragraphs[0].StyleId);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Run_StyleId_IsNull_ByDefault()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Null(document.AddParagraph(string.Empty).AddRun("text").StyleId);
    }

    [Fact]
    public void Run_StyleId_ReadsAStoredReference()
    {
        byte[] package = OoxmlValidation.CreateDocxWithStyleReference("text", null, "Emphasis");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        Assert.Equal("Emphasis", document.Paragraphs[0].Runs[0].StyleId);
    }

    [Fact]
    public void Run_StyleId_Throws_WhenTheValueIsMissing()
    {
        byte[] package = OoxmlValidation.CreateDocxWithStoredRunStyleReference("text", null);

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        Assert.Throws<InvalidOfficeFileException>(() => document.Paragraphs[0].Runs[0].StyleId);
    }

    [Fact]
    public void Run_StyleId_IsInsertedInSchemaOrder()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun("text");

        // Applied in reverse schema order on purpose.
        run.Bold = true;
        run.FontFamily = "Verdana";
        run.StyleId = "Emphasis";
        byte[] produced = SaveBytes.ToBytes(document);

        W.RunProperties? properties = OoxmlValidation.ReadDocxFirstRunProperties(produced);
        Assert.NotNull(properties);
        Assert.IsType<W.RunStyle>(properties.ChildElements[0]);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Run_StyleId_OverwritesAnExistingReference()
    {
        byte[] package = OoxmlValidation.CreateDocxWithStyleReference("text", null, "Old");
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));
        Run run = document.Paragraphs[0].Runs[0];

        run.StyleId = "New";
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Equal("New", run.StyleId);
        Assert.Single(OoxmlValidation.ReadDocxFirstRunProperties(produced)!.Elements<W.RunStyle>());
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Run_StyleId_Null_ClearsTheReference()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun("text");
        run.StyleId = "Emphasis";

        run.StyleId = null;
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Null(run.StyleId);
        Assert.Null(OoxmlValidation.ReadDocxFirstRunProperties(produced));
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Run_StyleId_Null_KeepsOtherFormatting()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun("text");
        run.StyleId = "Emphasis";
        run.Bold = true;

        run.StyleId = null;

        Assert.Null(run.StyleId);
        Assert.True(run.Bold);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("\t")]
    [InlineData("\n")]
    [InlineData("\r")]
    [InlineData("a\tb")]
    [InlineData("a\nb")]
    public void Run_StyleId_RejectsAValueThatIsNotAnId(string value)
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun("text");

        Assert.Throws<ArgumentException>(() => run.StyleId = value);
    }

    [Fact]
    public void Run_StyleId_RejectsAValueLongerThanTheSchemaAllows()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun("text");

        Assert.Throws<ArgumentException>(() => run.StyleId = new string('a', 254));
    }

    [Fact]
    public void Run_StyleId_ARejectedValueLeavesTheMarkupUntouched()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun("text");
        run.StyleId = "Emphasis";

        Assert.Throws<ArgumentException>(() => run.StyleId = "x\u0000");

        Assert.Equal("Emphasis", run.StyleId);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void Run_StyleId_RoundTripsThroughSaveAndOpen()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph(string.Empty).AddRun("text").StyleId = "Emphasis";
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Open(stream);
        Assert.Equal("Emphasis", reopened.Paragraphs[0].Runs[0].StyleId);
    }

    [Fact]
    public void ParagraphAndRun_StyleIds_CoexistInOneParagraph()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        paragraph.StyleId = "Heading1";
        paragraph.AddRun("text").StyleId = "Emphasis";
        byte[] produced = SaveBytes.ToBytes(document);

        W.Paragraph? first = OoxmlValidation.ReadDocxFirstParagraph(produced);
        Assert.Equal("Heading1", first?.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Equal("Emphasis", first?.Descendants<W.RunStyle>().FirstOrDefault()?.Val?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }
}
