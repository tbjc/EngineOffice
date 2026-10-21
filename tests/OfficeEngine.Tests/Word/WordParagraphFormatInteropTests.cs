using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Interoperability checks for the paragraph formatting added in Phase 2, increment 2.2. Every
/// assertion is made against the Open XML SDK or against the raw package content, never through the
/// code under test.
/// </summary>
public sealed class WordParagraphFormatInteropTests
{
    [Fact]
    public void ProducedDocument_WithParagraphFormatting_IsSchemaValid()
    {
        byte[] bytes = CreateDocumentWithParagraphFormatting();

        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void ProducedDocument_ParagraphFormatting_IsReadableByTheOpenXmlSdk()
    {
        byte[] bytes = CreateDocumentWithParagraphFormatting();

        W.ParagraphProperties? properties = OoxmlValidation.ReadDocxFirstParagraphProperties(bytes);

        Assert.NotNull(properties);
        Assert.Equal(W.JustificationValues.Center, properties.Justification?.Val?.Value);
        Assert.Equal("240", properties.Indentation?.FirstLine?.Value);
        Assert.Equal("480", properties.Indentation?.Left?.Value);
        Assert.Equal("120", properties.Indentation?.Hanging?.Value);
        Assert.Equal("120", properties.SpacingBetweenLines?.Before?.Value);
        Assert.Equal("240", properties.SpacingBetweenLines?.After?.Value);
        Assert.Equal("360", properties.SpacingBetweenLines?.Line?.Value);
        Assert.Equal(W.LineSpacingRuleValues.Exact, properties.SpacingBetweenLines?.LineRule?.Value);
    }

    [Fact]
    public void ProducedDocument_WritesParagraphPropertiesInSchemaOrder()
    {
        // The formatting is applied in the opposite order to the one CT_PPr requires, so the test
        // fails if the library appends properties instead of inserting them in schema order.
        byte[] bytes = CreateDocumentWithParagraphFormatting();

        string xml = TestPackage.ReadEntryText(bytes, "word/document.xml");
        int justification = xml.IndexOf("<w:jc", StringComparison.Ordinal);
        int indentation = xml.IndexOf("<w:ind", StringComparison.Ordinal);
        int spacing = xml.IndexOf("<w:spacing", StringComparison.Ordinal);

        Assert.True(
            spacing >= 0 && indentation > spacing && justification > indentation,
            $"Paragraph properties are not in schema order (spacing={spacing}, ind={indentation}, jc={justification}).");
    }

    [Fact]
    public void OpenedDocument_WithMixedParagraphFormatting_PreservesUnmodelledMarkup()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithMixedParagraphFormatting("mixto");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Paragraph paragraph = document.Paragraphs[0];

        Assert.Equal(ParagraphAlignment.Center, paragraph.Alignment);
        Assert.Equal(12.0, paragraph.FirstLineIndentPoints);
        Assert.Equal(18.0, paragraph.LeftIndentPoints);
        Assert.Equal(6.0, paragraph.SpacingBeforePoints);
        Assert.Null(paragraph.RightIndentPoints);

        paragraph.Alignment = ParagraphAlignment.Right;
        paragraph.LeftIndentPoints = 24.0;

        byte[] produced = SaveBytes.ToBytes(document);
        string xml = TestPackage.ReadEntryText(produced, "word/document.xml");

        Assert.Contains("w:keepNext", xml, StringComparison.Ordinal);
        Assert.Equal(1, xml.Split("<w:jc", StringSplitOptions.None).Length - 1);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        W.ParagraphProperties? properties = OoxmlValidation.ReadDocxFirstParagraphProperties(produced);
        Assert.Equal(W.JustificationValues.Right, properties?.Justification?.Val?.Value);
        Assert.Equal("480", properties?.Indentation?.Left?.Value);
        Assert.Equal("240", properties?.Indentation?.FirstLine?.Value);
    }

    [Fact]
    public void OpenedDocument_WithNumbering_KeepsItAheadOfTheModelledProperties()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithNumbering("numerado");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Paragraph paragraph = document.Paragraphs[0];
        paragraph.Alignment = ParagraphAlignment.Right;
        paragraph.LeftIndentPoints = 12.0;

        byte[] produced = SaveBytes.ToBytes(document);
        string xml = TestPackage.ReadEntryText(produced, "word/document.xml");
        int numbering = xml.IndexOf("<w:numPr", StringComparison.Ordinal);
        int indentation = xml.IndexOf("<w:ind", StringComparison.Ordinal);
        int justification = xml.IndexOf("<w:jc", StringComparison.Ordinal);

        Assert.Contains("w:numPr", xml, StringComparison.Ordinal);
        Assert.True(
            numbering >= 0 && indentation > numbering && justification > indentation,
            $"w:numPr must precede the properties added by the library (numPr={numbering}, ind={indentation}, jc={justification}).");
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void OpenedDocument_WithUnmodelledParagraphFlags_KeepsSchemaOrder()
    {
        // w:overflowPunct, w:bidi, w:adjustRightInd and w:snapToGrid must all stay ahead of the
        // properties this version writes (w:spacing, w:ind, w:jc), or the package becomes invalid.
        byte[] fixture = OoxmlValidation.CreateDocxWithUnmodelledParagraphFlags("bidi");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Paragraph paragraph = document.Paragraphs[0];
        paragraph.SpacingBeforePoints = 6.0;
        paragraph.LeftIndentPoints = 12.0;
        paragraph.Alignment = ParagraphAlignment.Center;

        byte[] produced = SaveBytes.ToBytes(document);
        string xml = TestPackage.ReadEntryText(produced, "word/document.xml");

        int punctuation = xml.IndexOf("<w:overflowPunct", StringComparison.Ordinal);
        int bidi = xml.IndexOf("<w:bidi", StringComparison.Ordinal);
        int rightIndent = xml.IndexOf("<w:adjustRightInd", StringComparison.Ordinal);
        int snapToGrid = xml.IndexOf("<w:snapToGrid", StringComparison.Ordinal);
        int spacing = xml.IndexOf("<w:spacing", StringComparison.Ordinal);
        int indentation = xml.IndexOf("<w:ind", StringComparison.Ordinal);
        int justification = xml.IndexOf("<w:jc", StringComparison.Ordinal);

        Assert.True(
            punctuation >= 0 && bidi > punctuation && rightIndent > bidi && snapToGrid > rightIndent
                && spacing > snapToGrid && indentation > spacing && justification > indentation,
            $"Unmodelled flags are not preserved in schema order (overflowPunct={punctuation}, bidi={bidi}, adjustRightInd={rightIndent}, snapToGrid={snapToGrid}, spacing={spacing}, ind={indentation}, jc={justification}).");

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    private static byte[] CreateDocumentWithParagraphFormatting()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("Formato de párrafo");

        // Deliberately applied in reverse schema order (jc, ind, spacing).
        paragraph.Alignment = ParagraphAlignment.Center;
        paragraph.FirstLineIndentPoints = 12.0;
        paragraph.LeftIndentPoints = 24.0;
        paragraph.HangingIndentPoints = 6.0;
        paragraph.SpacingBeforePoints = 6.0;
        paragraph.SpacingAfterPoints = 12.0;
        paragraph.LineSpacingPoints = 18.0;

        return SaveBytes.ToBytes(document);
    }
}
