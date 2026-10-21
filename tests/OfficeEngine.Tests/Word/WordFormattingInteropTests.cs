using DocumentFormat.OpenXml.Validation;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Interoperability checks for Phase 2, increment 1. Every assertion is made against the
/// Open XML SDK or against the raw package content, never through the code under test.
/// </summary>
public sealed class WordFormattingInteropTests
{
    [Fact]
    public void ProducedDocument_WithCharacterFormatting_IsSchemaValid()
    {
        byte[] bytes = CreateDocumentWithCharacterFormatting();

        IReadOnlyList<ValidationErrorInfo> errors = OoxmlValidation.ValidateDocx(bytes);

        Assert.Empty(errors);
    }

    [Fact]
    public void ProducedDocument_CharacterFormatting_IsReadableByTheOpenXmlSdk()
    {
        byte[] bytes = CreateDocumentWithCharacterFormatting();

        W.RunProperties? properties = OoxmlValidation.ReadDocxFirstRunProperties(bytes);

        Assert.NotNull(properties);
        Assert.NotNull(properties.Bold);
        Assert.Null(properties.Bold.Val);
        Assert.NotNull(properties.Italic);
        Assert.Null(properties.Italic.Val);
        Assert.Equal("Georgia", properties.RunFonts?.Ascii?.Value);
        Assert.Equal("28", properties.FontSize?.Val?.Value);
        Assert.Equal("112233", properties.Color?.Val?.Value);
        Assert.Equal<W.UnderlineValues?>(W.UnderlineValues.Double, properties.Underline?.Val?.Value);
    }

    [Fact]
    public void ProducedDocument_ExplicitlyDisabledBold_IsReadableAsDisabledByTheSdk()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph(string.Empty).AddRun("x").Bold = false;
        byte[] bytes = SaveBytes.ToBytes(document);

        W.RunProperties? properties = OoxmlValidation.ReadDocxFirstRunProperties(bytes);

        Assert.NotNull(properties?.Bold);
        Assert.Equal<bool?>(false, properties.Bold.Val?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void ProducedDocument_WritesRunPropertiesInSchemaOrder()
    {
        // Formatting is applied in reverse schema order on purpose, so that producing valid
        // output depends on the insertion logic rather than on the order of the assignments.
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun("orden");
        run.Underline = UnderlineStyle.SingleLine;
        run.FontSizePoints = 12;
        run.Color = RgbColor.FromRgb(0x01, 0x02, 0x03);
        run.Italic = true;
        run.Bold = true;
        run.FontFamily = "Verdana";
        byte[] bytes = SaveBytes.ToBytes(document);

        string documentXml = TestPackage.ReadEntryText(bytes, "word/document.xml");

        int start = documentXml.IndexOf("<w:rPr>", StringComparison.Ordinal);
        Assert.True(start >= 0, "The run must carry a w:rPr element.");
        int end = documentXml.IndexOf("</w:rPr>", start, StringComparison.Ordinal);
        Assert.True(end > start, "The w:rPr element must be closed.");
        string properties = documentXml.Substring(start, end - start);

        int fonts = properties.IndexOf("<w:rFonts", StringComparison.Ordinal);
        int bold = properties.IndexOf("<w:b", StringComparison.Ordinal);
        int italic = properties.IndexOf("<w:i", StringComparison.Ordinal);
        int color = properties.IndexOf("<w:color", StringComparison.Ordinal);
        int size = properties.IndexOf("<w:sz", StringComparison.Ordinal);
        int underline = properties.IndexOf("<w:u", StringComparison.Ordinal);

        Assert.True(fonts >= 0, "w:rFonts was not written.");
        Assert.True(bold >= 0, "w:b was not written.");
        Assert.True(italic >= 0, "w:i was not written.");
        Assert.True(color >= 0, "w:color was not written.");
        Assert.True(size >= 0, "w:sz was not written.");
        Assert.True(underline >= 0, "w:u was not written.");

        Assert.True(
            fonts < bold && bold < italic && italic < color && color < size && size < underline,
            $"Run properties are not in CT_RPr schema order: {properties}");

        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void OpenedDocument_WithRunFormatting_ReportsItThroughTheApi()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithFormattedRun("externo");

        // The fixture must itself be schema-valid, otherwise the API is not being tested fairly.
        Assert.Empty(OoxmlValidation.ValidateDocx(fixture));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Run run = Assert.Single(document.Paragraphs[0].Runs);

        Assert.Equal("externo", run.Text);
        Assert.Equal("Georgia", run.FontFamily);
        Assert.Equal<double?>(14d, run.FontSizePoints);
        Assert.Equal<bool?>(true, run.Bold);
        Assert.Equal<bool?>(true, run.Italic);
        Assert.Equal<UnderlineStyle?>(UnderlineStyle.DoubleLine, run.Underline);
        Assert.Equal<RgbColor?>(RgbColor.FromRgb(0x00, 0xFF, 0x00), run.Color);
    }

    [Fact]
    public void OpenedDocument_UnmodelledRunProperties_SurviveEditingModelledOnes()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithFormattedRun("externo");
        byte[] produced;

        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Run run = document.Paragraphs[0].Runs[0];
            run.Bold = false;
            run.Color = RgbColor.FromRgb(0x01, 0x02, 0x03);
            produced = SaveBytes.ToBytes(document);
        }

        string documentXml = TestPackage.ReadEntryText(produced, "word/document.xml");

        // w:highlight and the font name were not edited, so both must still be present.
        Assert.Contains("w:highlight", documentXml, StringComparison.Ordinal);
        Assert.Contains("Georgia", documentXml, StringComparison.Ordinal);
        Assert.Contains("010203", documentXml, StringComparison.Ordinal);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void OpenedDocument_WithThemeColor_ReportsNoExplicitColourAndPreservesMarkup()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithThemeColor("tema");

        // The fixture must itself be schema-valid, otherwise the check below proves nothing.
        Assert.Empty(OoxmlValidation.ValidateDocx(fixture));

        byte[] produced;

        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Run run = document.Paragraphs[0].Runs[0];

            // A theme colour has no fixed RGB value, so it must not be guessed.
            Assert.Null(run.Color);

            run.Bold = true;
            produced = SaveBytes.ToBytes(document);
        }

        string documentXml = TestPackage.ReadEntryText(produced, "word/document.xml");
        Assert.Contains("themeColor", documentXml, StringComparison.Ordinal);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void OpenedDocument_WithMalformedFontSize_ThrowsInvalidOfficeFileException()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithMalformedFontSize("malo");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Run run = document.Paragraphs[0].Runs[0];

        // A size that cannot be read must fail loudly instead of silently reporting no size.
        Assert.Throws<InvalidOfficeFileException>(() => _ = run.FontSizePoints);
    }

    [Fact]
    public void OpenedDocument_WithMalformedOnOffValue_ThrowsInvalidOfficeFileException()
    {
        // The SDK will not author an invalid ST_OnOff value, so the whole document part is
        // rewritten as raw XML over an otherwise valid package.
        byte[] valid = OoxmlValidation.CreateDocxWithFormattedRun("raro");
        const string malformedDocumentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:body><w:p><w:r><w:rPr><w:b w:val=\"bogus\"/></w:rPr><w:t>raro</w:t></w:r></w:p></w:body>"
            + "</w:document>";
        byte[] fixture = TestPackage.WithEntryText(valid, "word/document.xml", malformedDocumentXml);

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Run run = document.Paragraphs[0].Runs[0];

        // The framework's FormatException must not leak: a malformed file is a typed OfficeEngine error.
        Assert.Throws<InvalidOfficeFileException>(() => _ = run.Bold);
    }

    [Fact]
    public void OpenedDocument_WithUnknownUnderlineStyle_ThrowsInvalidOfficeFileException()
    {
        // The SDK refuses to author an unknown ST_Underline value, so it is injected into a
        // valid package as raw XML instead.
        byte[] valid = OoxmlValidation.CreateDocxWithFormattedRun("raro");
        byte[] fixture = TestPackage.WithEntryText(
            valid,
            "word/document.xml",
            TestPackage.ReadEntryText(valid, "word/document.xml")
                .Replace("w:val=\"double\"", "w:val=\"bogus\"", StringComparison.Ordinal));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Run run = document.Paragraphs[0].Runs[0];

        // An underline value outside the modelled set must fail loudly, not map to a wrong style.
        Assert.Throws<InvalidOfficeFileException>(() => _ = run.Underline);
    }

    /// <summary>
    /// Creates a document whose run formatting is applied deliberately out of schema order.
    /// </summary>
    private static byte[] CreateDocumentWithCharacterFormatting()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun("Verificación");
        run.Underline = UnderlineStyle.DoubleLine;
        run.FontSizePoints = 14;
        run.Color = RgbColor.FromRgb(0x11, 0x22, 0x33);
        run.Italic = true;
        run.Bold = true;
        run.FontFamily = "Georgia";
        return SaveBytes.ToBytes(document);
    }
}
