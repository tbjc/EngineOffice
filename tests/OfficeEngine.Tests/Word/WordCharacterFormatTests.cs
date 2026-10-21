using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Phase 2, increment 1: runs and character formatting (font family, size, bold, italic,
/// underline and colour).
/// </summary>
public sealed class WordCharacterFormatTests
{
    [Fact]
    public void CreateParagraph_ExposesASingleRunWithNoExplicitFormatting()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);

        Assert.Empty(paragraph.Runs);

        Run run = paragraph.AddRun("hola");
        Assert.Equal("hola", run.Text);

        // Tri-state: nothing is set, so everything inherits from styles or defaults.
        Assert.Null(run.FontFamily);
        Assert.Null(run.FontSizePoints);
        Assert.Null(run.Bold);
        Assert.Null(run.Italic);
        Assert.Null(run.Underline);
        Assert.Null(run.Color);
    }

    [Fact]
    public void AddParagraph_WithText_StillCreatesASingleRun()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("una línea");

        Run run = Assert.Single(paragraph.Runs);
        Assert.Equal("una línea", run.Text);
        Assert.Null(run.Bold);
    }

    [Fact]
    public void AddRun_AppendsRunsAndPreservesTheirOrder()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);

        paragraph.AddRun("uno ");
        paragraph.AddRun("dos ");
        paragraph.AddRun("tres");

        Assert.Equal(3, paragraph.Runs.Count);
        Assert.Equal("uno ", paragraph.Runs[0].Text);
        Assert.Equal("dos ", paragraph.Runs[1].Text);
        Assert.Equal("tres", paragraph.Runs[2].Text);
        Assert.Equal("uno dos tres", paragraph.Text);
    }

    [Fact]
    public void AppendText_StillCreatesAVisibleRun()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("a");

        paragraph.AppendText("b");

        Assert.Equal(2, paragraph.Runs.Count);
        Assert.Equal("ab", paragraph.Text);
    }

    [Fact]
    public void Run_Text_CanBeReplaced_WithoutLosingCharacterFormatting()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun("old");
        run.Bold = true;

        run.Text = "new";

        Assert.Equal("new", run.Text);
        Assert.Equal<bool?>(true, run.Bold);
    }

    [Fact]
    public void Run_CharacterFormatting_RoundTripsThroughSaveAndOpen()
    {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create())
        {
            Run run = document.AddParagraph(string.Empty).AddRun("Formateado");
            run.FontFamily = "Georgia";
            run.FontSizePoints = 14;
            run.Bold = true;
            run.Italic = true;
            run.Underline = UnderlineStyle.DoubleLine;
            run.Color = RgbColor.FromRgb(0x11, 0x22, 0x33);

            bytes = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(bytes));

        Run reopenedRun = Assert.Single(reopened.Paragraphs[0].Runs);
        Assert.Equal("Formateado", reopenedRun.Text);
        Assert.Equal("Georgia", reopenedRun.FontFamily);
        Assert.Equal<double?>(14d, reopenedRun.FontSizePoints);
        Assert.Equal<bool?>(true, reopenedRun.Bold);
        Assert.Equal<bool?>(true, reopenedRun.Italic);
        Assert.Equal<UnderlineStyle?>(UnderlineStyle.DoubleLine, reopenedRun.Underline);
        Assert.Equal<RgbColor?>(RgbColor.FromRgb(0x11, 0x22, 0x33), reopenedRun.Color);
    }

    [Fact]
    public void Run_HalfPointFontSize_RoundTripsWithoutRounding()
    {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create())
        {
            Run run = document.AddParagraph(string.Empty).AddRun("medio punto");
            run.FontSizePoints = 11.5;
            bytes = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(bytes));
        Assert.Equal<double?>(11.5, reopened.Paragraphs[0].Runs[0].FontSizePoints);
    }

    [Fact]
    public void Run_TriStateFormatting_DistinguishesUnsetFromExplicitlyDisabled()
    {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create())
        {
            Run run = document.AddParagraph(string.Empty).AddRun("x");

            Assert.Null(run.Bold);

            run.Bold = true;
            run.Bold = false;
            Assert.Equal<bool?>(false, run.Bold);

            bytes = SaveBytes.ToBytes(document);
        }

        // The distinction must survive the file: false is written, not dropped.
        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(bytes));
        Run reopenedRun = reopened.Paragraphs[0].Runs[0];
        Assert.Equal<bool?>(false, reopenedRun.Bold);

        // Clearing it removes the property entirely so the run inherits again.
        reopenedRun.Bold = null;
        Assert.Null(reopenedRun.Bold);

        using WordDocument third = WordDocument.Open(TestPackage.ToStream(SaveBytes.ToBytes(reopened)));
        Assert.Null(third.Paragraphs[0].Runs[0].Bold);
    }

    [Fact]
    public void Run_Underline_NoneIsDifferentFromUnset()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun("subrayado");

        Assert.Null(run.Underline);

        run.Underline = UnderlineStyle.None;
        Assert.Equal<UnderlineStyle?>(UnderlineStyle.None, run.Underline);

        run.Underline = null;
        Assert.Null(run.Underline);
    }

    [Fact]
    public void Run_Underline_RejectsAValueOutsideTheModelledSetWithoutTouchingTheDocument()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun("x");

        // Only possible by casting, but a value outside the enum must not be written as a wrong style.
        Assert.Throws<ArgumentOutOfRangeException>(() => run.Underline = (UnderlineStyle)999);

        // The rejected assignment must not have created any partial markup either.
        Assert.Null(run.Underline);
        byte[] bytes = SaveBytes.ToBytes(document);
        string documentXml = TestPackage.ReadEntryText(bytes, "word/document.xml");
        Assert.DoesNotContain("<w:u", documentXml, StringComparison.Ordinal);
        Assert.Empty(OoxmlValidation.ValidateDocx(bytes));
    }

    [Fact]
    public void Run_AllUnderlineStyles_RoundTrip()
    {
        UnderlineStyle[] styles =
        [
            UnderlineStyle.None, UnderlineStyle.SingleLine, UnderlineStyle.Words,
            UnderlineStyle.DoubleLine, UnderlineStyle.Thick, UnderlineStyle.Dotted,
            UnderlineStyle.DottedHeavy, UnderlineStyle.Dash, UnderlineStyle.DashedHeavy,
            UnderlineStyle.DashLong, UnderlineStyle.DashLongHeavy, UnderlineStyle.DotDash,
            UnderlineStyle.DashDotHeavy, UnderlineStyle.DotDotDash, UnderlineStyle.DashDotDotHeavy,
            UnderlineStyle.Wave, UnderlineStyle.WavyHeavy, UnderlineStyle.WavyDouble,
        ];

        byte[] bytes;
        using (WordDocument document = WordDocument.Create())
        {
            Paragraph paragraph = document.AddParagraph(string.Empty);
            for (int i = 0; i < styles.Length; i++)
            {
                Run run = paragraph.AddRun($"u{i}");
                run.Underline = styles[i];
            }

            bytes = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(bytes));
        Assert.Equal(styles, reopened.Paragraphs[0].Runs.Select(r => r.Underline!.Value));
    }

    [Fact]
    public void Run_CharacterFormatting_RoundTripsThroughAFile()
    {
        string directory = TestData.NewTemporaryDirectory();
        string path = Path.Combine(directory, "formateado.docx");
        try
        {
            using (WordDocument document = WordDocument.Create())
            {
                Run run = document.AddParagraph(string.Empty).AddRun("en disco");
                run.Bold = true;
                run.FontFamily = "Verdana";
                run.Color = RgbColor.FromRgb(0xFF, 0x00, 0x00);
                document.Save(path);
            }

            using WordDocument reopened = WordDocument.Open(path);
            Run reopenedRun = Assert.Single(reopened.Paragraphs[0].Runs);
            Assert.Equal<bool?>(true, reopenedRun.Bold);
            Assert.Equal("Verdana", reopenedRun.FontFamily);
            Assert.Equal<RgbColor?>(RgbColor.FromRgb(0xFF, 0x00, 0x00), reopenedRun.Color);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Paragraph_TextSetter_PreservesParagraphPropertiesAndDropsTheOldRuns()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithCenteredParagraph("antes");
        byte[] produced;

        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Paragraph paragraph = document.Paragraphs[0];
            Assert.Equal("antes", paragraph.Text);

            paragraph.Text = "después";
            produced = SaveBytes.ToBytes(document);
        }

        string documentXml = TestPackage.ReadEntryText(produced, "word/document.xml");
        Assert.Contains("después", documentXml, StringComparison.Ordinal);
        Assert.DoesNotContain("antes", documentXml, StringComparison.Ordinal);

        // The paragraph property this version does not model must not be dropped by a text edit.
        Assert.Contains("w:jc", documentXml, StringComparison.Ordinal);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Run_RejectsInvalidArguments()
    {
        using WordDocument document = WordDocument.Create();
        Run run = document.AddParagraph(string.Empty).AddRun("x");

        Assert.Throws<ArgumentNullException>(() => run.Text = null!);
        Assert.Throws<ArgumentNullException>(() => document.AddParagraph(null!));
        Assert.Throws<ArgumentNullException>(() => document.Paragraphs[0].AddRun(null!));

        Assert.Throws<ArgumentException>(() => run.FontFamily = string.Empty);
        Assert.Throws<ArgumentException>(() => run.FontFamily = "   ");

        Assert.Throws<ArgumentOutOfRangeException>(() => run.FontSizePoints = 0);
        Assert.Throws<ArgumentOutOfRangeException>(() => run.FontSizePoints = -12);
        Assert.Throws<ArgumentOutOfRangeException>(() => run.FontSizePoints = double.NaN);
        Assert.Throws<ArgumentOutOfRangeException>(() => run.FontSizePoints = 11.3);

        // ...and the value that was rejected must not have been written.
        Assert.Null(run.FontSizePoints);
    }

    [Fact]
    public void Run_ClearingEachProperty_RestoresInheritance()
    {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create())
        {
            Run run = document.AddParagraph(string.Empty).AddRun("x");
            run.FontFamily = "Arial";
            run.FontSizePoints = 12;
            run.Bold = true;
            run.Italic = true;
            run.Underline = UnderlineStyle.Wave;
            run.Color = RgbColor.FromRgb(1, 2, 3);

            run.FontFamily = null;
            run.FontSizePoints = null;
            run.Italic = null;
            run.Color = null;

            Assert.Null(run.FontFamily);
            Assert.Null(run.FontSizePoints);
            Assert.Null(run.Italic);
            Assert.Null(run.Color);

            // The properties that were not cleared must survive.
            Assert.Equal<bool?>(true, run.Bold);
            Assert.Equal<UnderlineStyle?>(UnderlineStyle.Wave, run.Underline);

            bytes = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(bytes));
        Run reopenedRun = reopened.Paragraphs[0].Runs[0];
        Assert.Null(reopenedRun.FontFamily);
        Assert.Null(reopenedRun.FontSizePoints);
        Assert.Null(reopenedRun.Italic);
        Assert.Null(reopenedRun.Color);
        Assert.Equal<bool?>(true, reopenedRun.Bold);
        Assert.Equal<UnderlineStyle?>(UnderlineStyle.Wave, reopenedRun.Underline);
    }

    [Fact]
    public void Run_ExplicitlyDisabledItalic_RoundTrips()
    {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph(string.Empty).AddRun("x").Italic = false;
            bytes = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(bytes));
        Assert.Equal<bool?>(false, reopened.Paragraphs[0].Runs[0].Italic);
    }

    [Fact]
    public void RunAndParagraph_ToString_ReturnTheText()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        Run run = paragraph.AddRun("resumen");

        Assert.Equal("resumen", run.ToString());
        Assert.Equal("resumen", paragraph.ToString());
    }

    [Fact]
    public void AppendText_WithEmptyString_IsIgnored()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("base");

        paragraph.AppendText(string.Empty);

        Assert.Single(paragraph.Runs);
        Assert.Equal("base", paragraph.Text);
    }

    [Fact]
    public void Runs_AreLiveViewsOverTheSameContent()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph(string.Empty);
        Run first = paragraph.AddRun("a");

        // A fresh snapshot sees the same run, and editing through either handle is visible.
        Run second = Assert.Single(paragraph.Runs);
        second.Text = "b";

        Assert.Equal("b", first.Text);
        Assert.Equal("b", paragraph.Text);
        Assert.Same(first, first);
    }
}
