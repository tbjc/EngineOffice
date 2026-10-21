using System;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.6: interoperability of sections and page setup, verified with the Open XML SDK instead
/// of the code under test — produced packages are read back and validated independently.
/// </summary>
public sealed class WordSectionInteropTests
{
    [Fact]
    public void ProducedDocument_WithPageSetup_IsSchemaValidAndReadable()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];
        section.PageSize = PageSize.A4.Rotate();
        section.Margins = new PageMargins(36.0, 54.0, 36.0, 54.0);
        section.StartType = SectionStartType.EvenPage;

        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        W.SectionProperties stored = Assert.Single(OoxmlValidation.ReadDocxSections(produced));
        Assert.Equal(
            ["type", "pgSz", "pgMar"],
            stored.ChildElements.Select(child => child.LocalName));

        W.PageSize? pageSize = stored.GetFirstChild<W.PageSize>();
        W.PageMargin? pageMargin = stored.GetFirstChild<W.PageMargin>();
        W.SectionType? sectionType = stored.GetFirstChild<W.SectionType>();

        Assert.Equal("16838", pageSize?.Width?.InnerText);
        Assert.Equal("11906", pageSize?.Height?.InnerText);
        Assert.Equal(W.PageOrientationValues.Landscape, pageSize?.Orient?.Value);
        Assert.Equal("720", pageMargin?.Top?.InnerText);
        Assert.Equal("1080", pageMargin?.Right?.InnerText);
        Assert.Equal(W.SectionMarkValues.EvenPage, sectionType?.Val?.Value);
    }

    [Fact]
    public void ProducedDocument_WithTwoSections_IsSchemaValidAndKeepsTheFinalSectionLast()
    {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].PageSize = PageSize.A4;
        document.AddParagraph("first");
        document.AddSection(SectionStartType.Continuous);
        document.AddParagraph("second");

        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(
            ["p", "p", "p", "sectPr"],
            OoxmlValidation.ReadDocxBodyChildLocalNames(produced));

        var sections = OoxmlValidation.ReadDocxSections(produced);
        Assert.Equal(2, sections.Count);
        Assert.Equal("11906", sections[0].GetFirstChild<W.PageSize>()?.Width?.InnerText);
        Assert.Equal(W.SectionMarkValues.Continuous, sections[1].GetFirstChild<W.SectionType>()?.Val?.Value);
        Assert.Equal("11906", sections[1].GetFirstChild<W.PageSize>()?.Width?.InnerText);
    }

    [Fact]
    public void OpenedDocument_WithUnmodelledSectionMarkup_KeepsItWhenPageSetupIsEdited()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(new W.SectionProperties(
            new W.PageSize { Width = 12240U, Height = 15840U },
            new W.PageMargin
            {
                Top = 1440,
                Right = 1440U,
                Bottom = 1440,
                Left = 1440U,
                Header = 360U,
                Footer = 360U,
                Gutter = 120U,
            },
            new W.Columns { Space = "720" },
            new W.TitlePage(),
            new W.DocGrid { Type = W.DocGridValues.Lines, LinePitch = 360 }));

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Section section = document.Sections[0];
            section.PageSize = PageSize.A4;
            section.Margins = new PageMargins(72.0, 72.0, 72.0, 72.0);
            section.StartType = SectionStartType.OddPage;
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        W.SectionProperties stored = Assert.Single(OoxmlValidation.ReadDocxSections(produced));
        W.PageSize? pageSize = stored.GetFirstChild<W.PageSize>();
        W.PageMargin? pageMargin = stored.GetFirstChild<W.PageMargin>();
        W.Columns? columns = stored.GetFirstChild<W.Columns>();
        W.DocGrid? docGrid = stored.GetFirstChild<W.DocGrid>();
        W.SectionType? sectionType = stored.GetFirstChild<W.SectionType>();

        Assert.Equal("11906", pageSize?.Width?.InnerText);
        Assert.Equal("1440", pageMargin?.Top?.InnerText);
        Assert.Equal("360", pageMargin?.Header?.InnerText);
        Assert.Equal("360", pageMargin?.Footer?.InnerText);
        Assert.Equal("120", pageMargin?.Gutter?.InnerText);
        Assert.Equal("720", columns?.Space?.Value);
        Assert.NotNull(stored.GetFirstChild<W.TitlePage>());
        Assert.Equal(W.DocGridValues.Lines, docGrid?.Type?.Value);
        Assert.Equal(360, docGrid?.LinePitch?.Value);
        Assert.Equal(W.SectionMarkValues.OddPage, sectionType?.Val?.Value);
    }

    [Fact]
    public void OpenedDocument_WithTwoSections_EditsOnlyTheRequestedOne()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(
            new W.Paragraph(
                new W.ParagraphProperties(
                    new W.SectionProperties(new W.PageSize { Width = 11906U, Height = 16838U })),
                new W.Run(new W.Text("first"))),
            new W.Paragraph(new W.Run(new W.Text("second"))),
            new W.SectionProperties(new W.PageSize { Width = 12240U, Height = 15840U }));

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            document.Sections[1].PageSize = PageSize.Letter.Rotate();
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        var sections = OoxmlValidation.ReadDocxSections(produced);
        Assert.Equal(2, sections.Count);
        Assert.Equal("11906", sections[0].GetFirstChild<W.PageSize>()?.Width?.InnerText);
        Assert.Null(sections[0].GetFirstChild<W.PageMargin>());
        Assert.Equal("15840", sections[1].GetFirstChild<W.PageSize>()?.Width?.InnerText);
        Assert.Equal("12240", sections[1].GetFirstChild<W.PageSize>()?.Height?.InnerText);
        Assert.Equal(
            W.PageOrientationValues.Landscape,
            sections[1].GetFirstChild<W.PageSize>()?.Orient?.Value);
    }

    [Fact]
    public void OpenedDocument_WithMalformedPageSize_ThrowsAndPreservesTheMarkup()
    {
        W.PageSize malformed = new() { Height = 15840U };
        OoxmlValidation.WithAttribute(malformed, "w", "wide");

        byte[] fixture = OoxmlValidation.CreateDocxWithBody(new W.SectionProperties(malformed));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Sections[0].PageSize);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Equal(
            "wide",
            OoxmlValidation.ReadDocxFinalSectionProperties(produced)?.GetFirstChild<W.PageSize>()?.Width?.InnerText);
    }

    [Fact]
    public void OpenedDocument_WithMalformedOrientation_Throws()
    {
        W.PageSize malformed = new() { Width = 12240U, Height = 15840U };
        OoxmlValidation.WithAttribute(malformed, "orient", "diagonal");

        byte[] fixture = OoxmlValidation.CreateDocxWithBody(new W.SectionProperties(malformed));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Sections[0].PageSize);
    }

    [Fact]
    public void OpenedDocument_WithMalformedPageMargin_Throws()
    {
        W.PageMargin malformed = new() { Right = 1440U, Bottom = 1440, Left = 1440U };
        OoxmlValidation.WithAttribute(malformed, "top", "wide");

        byte[] fixture = OoxmlValidation.CreateDocxWithBody(new W.SectionProperties(malformed));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Sections[0].Margins);
    }

    [Fact]
    public void OpenedDocument_WithSectionTypeWithoutValue_Throws()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(new W.SectionProperties(new W.SectionType()));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Sections[0].StartType);
    }

    [Fact]
    public void OpenedDocument_WithMalformedSectionType_Throws()
    {
        W.SectionType malformed = new();
        OoxmlValidation.WithAttribute(malformed, "val", "diagonal");

        byte[] fixture = OoxmlValidation.CreateDocxWithBody(new W.SectionProperties(malformed));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Sections[0].StartType);
    }
}
