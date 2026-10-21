using System;
using DocumentFormat.OpenXml.Validation;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.3 interoperability checks. Every assertion is made against the Open XML SDK or against
/// the raw package content, never through the code under test.
/// </summary>
public sealed class WordStyleInteropTests
{
    [Fact]
    public void ProducedDocument_WithStyleReferences_IsSchemaValid()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        paragraph.StyleId = "Heading1";
        paragraph.AddRun("text").StyleId = "Emphasis";
        byte[] bytes = SaveBytes.ToBytes(document);

        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void ProducedDocument_StyleReferences_AreReadableByTheOpenXmlSdk()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        paragraph.StyleId = "Heading1";
        paragraph.AddRun("text").StyleId = "Emphasis";
        byte[] bytes = SaveBytes.ToBytes(document);

        W.Paragraph? first = OoxmlValidation.ReadDocxFirstParagraph(bytes);

        Assert.Equal("Heading1", first?.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Equal("Emphasis", first?.Descendants<W.RunStyle>().FirstOrDefault()?.Val?.Value);
    }

    [Fact]
    public void ProducedDocument_DoesNotInventAStylePart()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("text").StyleId = "Heading1";
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.False(TestPackage.HasEntry(produced, "word/styles.xml"));
    }

    [Fact]
    public void ForeignDocument_StylesAreReadFromTheRawMarkup()
    {
        W.Style normal = new()
        {
            Type = W.StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true,
            StyleName = new W.StyleName { Val = "Normal" },
        };
        W.Style heading = new(
            new W.StyleName { Val = "Heading 1" },
            new W.BasedOn { Val = "Normal" },
            new W.StyleParagraphProperties(new W.SpacingBetweenLines { Before = "240" }),
            new W.StyleRunProperties(
                new W.RunFonts { Ascii = "Calibri Light", HighAnsi = "Calibri Light" },
                new W.Color { Val = "2F5496" },
                new W.FontSize { Val = "32" }))
        {
            Type = W.StyleValues.Paragraph,
            StyleId = "Heading1",
        };
        byte[] package = OoxmlValidation.CreateDocxWithStyles("body", normal, heading);

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));
        WordStyle style = document.FindStyle("Heading1")!;

        Assert.Equal("Heading 1", style.Name);
        Assert.Equal(WordStyleType.Paragraph, style.Type);
        Assert.False(style.IsDefault);
        Assert.Equal("Normal", style.BasedOnId);

        WordStyleFormatting effective = style.EffectiveFormatting;
        Assert.Equal("Calibri Light", effective.FontFamily);
        Assert.Equal(16.0, effective.FontSizePoints);
        Assert.Equal(12.0, effective.SpacingBeforePoints);
        Assert.Equal(RgbColor.FromRgb(0x2F, 0x54, 0x96), effective.Color);
        Assert.Null(effective.Bold);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void StylesPart_IsPreservedWhenTheBodyIsEdited()
    {
        W.Style custom = new(
            new W.StyleName { Val = "Custom" },
            new W.StyleRunProperties(
                new W.RunFonts { Ascii = "Georgia" },
                new W.Highlight { Val = W.HighlightColorValues.Yellow },
                new W.Underline { Val = W.UnderlineValues.Double }))
        {
            Type = W.StyleValues.Character,
            StyleId = "Custom",
        };
        byte[] package = OoxmlValidation.CreateDocxWithStyles("body", custom);
        string before = OoxmlValidation.ReadDocxStylesXml(package);

        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(package)))
        {
            document.Paragraphs[0].Text = "edited";
            document.Save(stream);
        }

        byte[] produced = stream.ToArray();
        Assert.Equal(before, OoxmlValidation.ReadDocxStylesXml(produced));
        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Equal("Custom", reopened.FindStyle("Custom")?.Name);
    }

    [Fact]
    public void DanglingStyleReference_RoundTrips()
    {
        byte[] package = OoxmlValidation.CreateDocxWithStoredParagraphStyleReference("text", "MissingStyle");
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(package)))
        {
            document.Save(stream);
        }

        byte[] produced = stream.ToArray();
        W.Paragraph? paragraph = OoxmlValidation.ReadDocxFirstParagraph(produced);
        Assert.Equal("MissingStyle", paragraph?.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void StyleReferences_WrittenByTheLibrary_SurviveAFileRoundTrip()
    {
        string path = Path.Combine(Path.GetTempPath(), $"oe-style-interop-{Guid.NewGuid():N}.docx");
        try
        {
            using (WordDocument document = WordDocument.Create())
            {
                Paragraph paragraph = document.AddParagraph(string.Empty);
                paragraph.StyleId = "Heading1";
                paragraph.AddRun("text").StyleId = "Emphasis";
                document.Save(path);
            }

            using WordDocument reopened = WordDocument.Open(path);
            Assert.Equal("Heading1", reopened.Paragraphs[0].StyleId);
            Assert.Equal("Emphasis", reopened.Paragraphs[0].Runs[0].StyleId);
            Assert.Equal("Emphasis", OoxmlValidation.ReadDocxFirstRun(File.ReadAllBytes(path))?.RunProperties?.RunStyle?.Val?.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
