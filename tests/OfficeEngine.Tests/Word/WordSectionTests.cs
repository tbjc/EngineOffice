using System;
using System.Linq;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.6: sections, page size/orientation and margins. Each property reads, writes and clears
/// its own markup, sections are reported in document order, and values that cannot be read fail with
/// a typed error instead of being guessed.
/// </summary>
public sealed class WordSectionTests
{
    [Fact]
    public void Create_HasOneLetterPortraitSectionWithOneInchMargins()
    {
        using WordDocument document = WordDocument.Create();

        Section section = Assert.Single(document.Sections);

        Assert.Equal(PageSize.Letter, section.PageSize);
        Assert.Equal(new PageMargins(72.0, 72.0, 72.0, 72.0), section.Margins);
        Assert.Null(section.StartType);
    }

    [Fact]
    public void Sections_ReportsEverySectionInDocumentOrder()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(
            new W.Paragraph(
                new W.ParagraphProperties(
                    new W.SectionProperties(new W.PageSize { Width = 11906U, Height = 16838U })),
                new W.Run(new W.Text("first"))),
            new W.Paragraph(new W.Run(new W.Text("second"))),
            new W.SectionProperties(new W.PageSize { Width = 12240U, Height = 15840U }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Equal(2, document.Sections.Count);
        Assert.Equal(PageSize.A4, document.Sections[0].PageSize);
        Assert.Equal(PageSize.Letter, document.Sections[1].PageSize);
    }

    [Fact]
    public void PageSize_CanBeSetReadAndCleared()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];

        section.PageSize = PageSize.A4.Rotate();
        Assert.Equal(PageSize.A4.Rotate(), section.PageSize);

        byte[] produced = SaveBytes.ToBytes(document);
        using (WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced)))
        {
            Assert.Equal(PageSize.A4.Rotate(), reopened.Sections[0].PageSize);
        }

        section.PageSize = null;
        Assert.Null(section.PageSize);

        byte[] cleared = SaveBytes.ToBytes(document);
        Assert.Null(OoxmlValidation.ReadDocxFinalSectionProperties(cleared)?.GetFirstChild<W.PageSize>());
        Assert.Empty(OoxmlValidation.ValidateDocx(cleared));
    }

    [Fact]
    public void PageSize_IsStoredInTwipsWithItsOrientation()
    {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].PageSize = PageSize.A4.Rotate();

        W.PageSize? stored = OoxmlValidation.ReadDocxFinalSectionProperties(SaveBytes.ToBytes(document))
            ?.GetFirstChild<W.PageSize>();

        Assert.Equal("16838", stored?.Width?.InnerText);
        Assert.Equal("11906", stored?.Height?.InnerText);
        Assert.Equal(W.PageOrientationValues.Landscape, stored?.Orient?.Value);
    }

    [Fact]
    public void PageSetup_RoundTripsThroughAFile()
    {
        string directory = TestData.NewTemporaryDirectory();
        string path = Path.Combine(directory, "sections.docx");
        try
        {
            using (WordDocument document = WordDocument.Create())
            {
                document.Sections[0].PageSize = PageSize.A4.Rotate();
                document.Sections[0].Margins = new PageMargins(36.0, 54.0, 36.0, 54.0);
                document.Save(path);
            }

            using WordDocument reopened = WordDocument.Open(path);

            Assert.Equal(PageSize.A4.Rotate(), reopened.Sections[0].PageSize);
            Assert.Equal(new PageMargins(36.0, 54.0, 36.0, 54.0), reopened.Sections[0].Margins);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void PageSize_RotateSwapsTheDimensionsAndTogglesTheOrientation()
    {
        PageSize rotated = PageSize.A4.Rotate();

        Assert.Equal(PageSize.A4.HeightPoints, rotated.WidthPoints);
        Assert.Equal(PageSize.A4.WidthPoints, rotated.HeightPoints);
        Assert.Equal(PageOrientation.Landscape, rotated.Orientation);
        Assert.Equal(PageSize.A4, rotated.Rotate());
    }

    [Fact]
    public void PageSize_RejectsInvalidValues()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageSize(0.0, 792.0, PageOrientation.Portrait));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageSize(0.01, 792.0, PageOrientation.Portrait));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageSize(-1.0, 792.0, PageOrientation.Portrait));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageSize(double.NaN, 792.0, PageOrientation.Portrait));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageSize(double.PositiveInfinity, 792.0, PageOrientation.Portrait));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageSize(double.MaxValue, 792.0, PageOrientation.Portrait));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageSize(612.0, 792.0, (PageOrientation)99));
    }

    [Fact]
    public void Margins_CanBeSetReadAndCleared()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];
        PageMargins margins = new(36.0, 54.0, 36.0, 54.0);

        section.Margins = margins;
        Assert.Equal(margins, section.Margins);

        byte[] produced = SaveBytes.ToBytes(document);
        using (WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced)))
        {
            Assert.Equal(margins, reopened.Sections[0].Margins);
        }

        section.Margins = null;
        Assert.Null(section.Margins);

        byte[] cleared = SaveBytes.ToBytes(document);
        Assert.Null(OoxmlValidation.ReadDocxFinalSectionProperties(cleared)?.GetFirstChild<W.PageMargin>());
        Assert.Empty(OoxmlValidation.ValidateDocx(cleared));
    }

    [Fact]
    public void Margins_AreStoredInTwips()
    {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].Margins = new PageMargins(36.0, 54.0, 18.0, 54.0);

        W.PageMargin? stored = OoxmlValidation.ReadDocxFinalSectionProperties(SaveBytes.ToBytes(document))
            ?.GetFirstChild<W.PageMargin>();

        Assert.Equal("720", stored?.Top?.InnerText);
        Assert.Equal("1080", stored?.Right?.InnerText);
        Assert.Equal("360", stored?.Bottom?.InnerText);
        Assert.Equal("1080", stored?.Left?.InnerText);
    }

    [Fact]
    public void Margins_AllowANegativeTopAndBottomButNotNegativeHorizontalMargins()
    {
        Assert.Equal(-36.0, new PageMargins(-36.0, 54.0, -18.0, 54.0).TopPoints);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageMargins(36.0, -1.0, 36.0, 36.0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PageMargins(36.0, 36.0, 36.0, -1.0));
    }

    [Fact]
    public void Margins_ReportAMissingRequiredMargin()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(new W.SectionProperties(
            new W.PageMargin { Top = 1440, Right = 1440U, Bottom = 1440 }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Sections[0].Margins);
    }

    [Fact]
    public void StartType_CanBeSetReadAndCleared()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];

        foreach (SectionStartType startType in Enum.GetValues<SectionStartType>())
        {
            section.StartType = startType;
            Assert.Equal(startType, section.StartType);

            byte[] produced = SaveBytes.ToBytes(document);
            using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
            Assert.Equal(startType, reopened.Sections[0].StartType);
        }

        section.StartType = null;
        Assert.Null(section.StartType);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void StartType_RejectsAnUndefinedValueWithoutTouchingTheDocument()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];

        Assert.Throws<ArgumentOutOfRangeException>(() => section.StartType = (SectionStartType)99);
        Assert.Null(section.StartType);
    }

    [Fact]
    public void AddSection_AppendsASecondSectionWithIndependentPageSetup()
    {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].PageSize = PageSize.A4;
        document.Sections[0].Margins = new PageMargins(72.0, 72.0, 72.0, 72.0);
        document.AddParagraph("first");

        Section second = document.AddSection(SectionStartType.Continuous);
        document.AddParagraph("second");
        second.PageSize = PageSize.Letter.Rotate();
        second.Margins = new PageMargins(36.0, 36.0, 36.0, 36.0);

        Assert.Equal(2, document.Sections.Count);
        Assert.Equal(PageSize.A4, document.Sections[0].PageSize);
        Assert.Equal(PageSize.Letter.Rotate(), document.Sections[1].PageSize);
        Assert.Equal(SectionStartType.Continuous, document.Sections[1].StartType);
        Assert.Equal("first\n\nsecond", document.Text);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Equal(2, reopened.Sections.Count);
        Assert.Equal(PageSize.A4, reopened.Sections[0].PageSize);
        Assert.Equal(new PageMargins(72.0, 72.0, 72.0, 72.0), reopened.Sections[0].Margins);
        Assert.Equal(PageSize.Letter.Rotate(), reopened.Sections[1].PageSize);
        Assert.Equal(new PageMargins(36.0, 36.0, 36.0, 36.0), reopened.Sections[1].Margins);
        Assert.Equal(SectionStartType.Continuous, reopened.Sections[1].StartType);
    }

    [Fact]
    public void AddSection_OnADocumentWithoutSectionProperties_CreatesTheDefaultSectionOnly()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(
            new W.Paragraph(new W.Run(new W.Text("text"))));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Empty(document.Sections);

        Section section = document.AddSection();

        Assert.Single(document.Sections);
        Assert.Equal(PageSize.Letter, section.PageSize);
        Assert.Null(section.StartType);
        Assert.Equal("text", document.Text);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void AddSection_RejectsAnUndefinedStartTypeWithoutTouchingTheDocument()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddSection((SectionStartType)99));
        Assert.Single(document.Sections);
        Assert.Empty(document.Paragraphs);
    }

    [Fact]
    public void Sections_ThrowsWhenDisposed()
    {
        WordDocument document = WordDocument.Create();
        document.Dispose();

        Assert.Throws<ObjectDisposedException>(() => document.Sections);
        Assert.Throws<ObjectDisposedException>(() => document.AddSection());
    }
}
